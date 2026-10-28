"""Evalúa las fórmulas del Excel de reparto lo bastante como para verificar el cierre."""

from __future__ import annotations

import re
from decimal import Decimal, ROUND_HALF_UP

TOKEN = re.compile(
    r"\s+|\"(?P<string>(?:[^\"]|\"\")*)\""
    r"|(?P<num>\d+(?:\.\d+)?)"
    r"|(?P<cmp><=|>=|<>)"
    r"|(?P<op>[&+\-*/=><])"
    r"|(?P<sep>[,()])"
    r"|(?P<ref>(?:[A-Za-z_][A-Za-z0-9_]*!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)"
    r"|(?P<name>[A-Za-z_][A-Za-z0-9_]*)"
)


def _tokens(formula: str) -> list[tuple[str, str]]:
    salida = []
    for match in TOKEN.finditer(formula):
        if match.group(0).strip() == "" and match.lastgroup is None:
            continue
        kind = match.lastgroup
        if kind is None:
            continue
        value = match.group(kind)
        if kind == "string":
            value = value.replace('""', '"')
        salida.append((kind, value))
    return salida


def _num(value) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, bool):
        return Decimal(int(value))
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def _truth(value) -> bool:
    if isinstance(value, bool):
        return value
    if value is None or value == "":
        return False
    if isinstance(value, (int, float, Decimal)) and Decimal(str(value)) == 0:
        return False
    return True


class Libro:
    def __init__(self, workbook):
        self.workbook = workbook
        self.memo: dict[tuple[str, str], object] = {}

    def celda(self, hoja: str, coord: str):
        coord = coord.replace("$", "")
        clave = (hoja, coord)
        if clave in self.memo:
            if self.memo[clave] == "PENDIENTE":
                raise RuntimeError(f"Referencia circular en {hoja}!{coord}")
            return self.memo[clave]
        self.memo[clave] = "PENDIENTE"
        valor = self.workbook[hoja][coord].value
        if isinstance(valor, str) and valor.startswith("="):
            resultado = _Eval(self, hoja, _tokens(valor[1:])).parse()
        else:
            resultado = valor
        self.memo[clave] = resultado
        return resultado


class _Eval:
    def __init__(self, libro: Libro, hoja: str, tokens: list[tuple[str, str]]):
        self.libro = libro
        self.hoja = hoja
        self.tokens = tokens
        self.indice = 0

    def parse(self):
        valor = self._concat()
        if self.indice != len(self.tokens):
            raise RuntimeError(f"Sobró {self.tokens[self.indice:]} en la fórmula")
        return valor

    def _peek(self):
        if self.indice >= len(self.tokens):
            return None
        return self.tokens[self.indice][1]

    def _pop(self):
        token = self.tokens[self.indice]
        self.indice += 1
        return token

    def _concat(self):
        izquierda = self._cmp()
        while self._peek() == "&":
            self._pop()
            derecha = self._cmp()
            izquierda = f"{'' if izquierda is None else izquierda}{'' if derecha is None else derecha}"
        return izquierda

    def _cmp(self):
        izquierda = self._suma()
        if self._peek() in {"=", ">", "<", ">=", "<=", "<>"}:
            operador = self._pop()[1]
            derecha = self._suma()
            return _comparar(operador, izquierda, derecha)
        return izquierda

    def _suma(self):
        izquierda = self._producto()
        while self._peek() in {"+", "-"}:
            operador = self._pop()[1]
            derecha = self._producto()
            izquierda = _num(izquierda) + _num(derecha) if operador == "+" else _num(izquierda) - _num(derecha)
        return izquierda

    def _producto(self):
        izquierda = self._unario()
        while self._peek() in {"*", "/"}:
            operador = self._pop()[1]
            derecha = self._unario()
            if operador == "*":
                izquierda = _num(izquierda) * _num(derecha)
            else:
                izquierda = _num(izquierda) / _num(derecha)
        return izquierda

    def _unario(self):
        if self._peek() == "-":
            self._pop()
            return -_num(self._unario())
        if self._peek() == "+":
            self._pop()
            return _num(self._unario())
        return self._primario()

    def _primario(self):
        kind, value = self._pop()
        if kind == "num":
            return Decimal(value)
        if kind == "string":
            return value
        if kind == "sep" and value == "(":
            adentro = self._concat()
            self._esperar(")")
            return adentro
        if kind == "name":
            self._esperar("(")
            return self._llamada(value)
        if kind == "ref":
            return self._resolver(value)
        raise RuntimeError(f"No entiendo {kind} {value}")

    def _llamada(self, nombre: str):
        nombre = nombre.upper()
        if nombre == "IF":
            condicion = self._concat()
            self._esperar(",")
            if _truth(condicion):
                elegido = self._concat()
                self._esperar(",")
                self._saltar()
                self._esperar(")")
                return elegido
            self._saltar()
            self._esperar(",")
            elegido = self._concat()
            self._esperar(")")
            return elegido
        if nombre in {"AND", "OR"}:
            valores = []
            while True:
                valores.append(_truth(self._concat()))
                if self._peek() == ",":
                    self._pop()
                    continue
                break
            self._esperar(")")
            return all(valores) if nombre == "AND" else any(valores)
        if nombre == "ROUND":
            numero = _num(self._concat())
            self._esperar(",")
            decimales = int(self._concat())
            self._esperar(")")
            cuantos = Decimal("1").scaleb(-decimales)
            return numero.quantize(cuantos, rounding=ROUND_HALF_UP)
        if nombre == "SUM":
            total = Decimal("0")
            while True:
                total += self._sumar_argumento()
                if self._peek() == ",":
                    self._pop()
                    continue
                break
            self._esperar(")")
            return total
        if nombre == "SUMIF":
            origen = self._rango()
            self._esperar(",")
            criterio = self._concat()
            self._esperar(",")
            suma = self._rango()
            self._esperar(")")
            return self._sumif(origen, criterio, suma)
        raise RuntimeError(f"Función no soportada: {nombre}")

    def _sumar_argumento(self) -> Decimal:
        if self.tokens[self.indice][0] == "ref" and ":" in self.tokens[self.indice][1]:
            return sum((_num(valor) for valor in self._valores(self._rango())), Decimal("0"))
        return _num(self._concat())

    def _rango(self) -> str:
        kind, value = self._pop()
        if kind != "ref" or ":" not in value:
            raise RuntimeError(f"Esperaba un rango y llegó {value}")
        return value

    def _valores(self, rango: str) -> list:
        hoja, coordenadas = self._partes(rango)
        inicio, fin = coordenadas.split(":")
        celdas = _expandir(inicio, fin)
        return [self.libro.celda(hoja, coord) for coord in celdas]

    def _sumif(self, origen: str, criterio, suma: str) -> Decimal:
        claves = self._valores(origen)
        valores = self._valores(suma)
        total = Decimal("0")
        for clave, valor in zip(claves, valores, strict=True):
            if _igual(clave, criterio):
                total += _num(valor)
        return total

    def _resolver(self, ref: str):
        if ":" in ref:
            raise RuntimeError(f"Rango suelto: {ref}")
        hoja, coord = self._partes(ref)
        return self.libro.celda(hoja, coord)

    def _partes(self, ref: str) -> tuple[str, str]:
        if "!" in ref:
            hoja, coord = ref.split("!", 1)
            return hoja, coord.replace("$", "")
        return self.hoja, ref.replace("$", "")

    def _esperar(self, simbolo: str) -> None:
        kind, value = self._pop()
        if value != simbolo:
            raise RuntimeError(f"Esperaba {simbolo} y llegó {value}")

    def _saltar(self) -> None:
        """Saltea un argumento sin evaluarlo, respetando paréntesis."""
        profundidad = 0
        while self.indice < len(self.tokens):
            value = self.tokens[self.indice][1]
            if value == "," and profundidad == 0:
                return
            if value == ")" and profundidad == 0:
                return
            if value == "(":
                profundidad += 1
            elif value == ")":
                profundidad -= 1
            self.indice += 1


def _igual(clave, criterio) -> bool:
    if clave is None:
        return criterio in ("", None)
    return str(clave).casefold() == str(criterio).casefold()


def _comparar(operador: str, izquierda, derecha) -> bool:
    if operador == "=":
        # En Excel 0 no es igual a una celda vacía.
        if _vacio(izquierda) or _vacio(derecha):
            return _vacio(izquierda) and _vacio(derecha)
        return _igual(izquierda, derecha) or _numeros_iguales(izquierda, derecha)
    if _vacio(izquierda) and not _vacio(derecha):
        izquierda = 0
    if _vacio(derecha) and not _vacio(izquierda):
        derecha = 0
    a, b = _num(izquierda), _num(derecha)
    return {
        ">": a > b,
        "<": a < b,
        ">=": a >= b,
        "<=": a <= b,
        "<>": a != b,
    }[operador]


def _vacio(valor) -> bool:
    return valor is None or valor == ""


def _numeros_iguales(izquierda, derecha) -> bool:
    if _vacio(izquierda) and _vacio(derecha):
        return True
    try:
        return _num(izquierda) == _num(derecha)
    except Exception:
        return False


def _expandir(inicio: str, fin: str) -> list[str]:
    c1, f1 = _coord(inicio)
    c2, f2 = _coord(fin)
    return [
        f"{_columna(columna)}{fila}"
        for columna in range(c1, c2 + 1)
        for fila in range(f1, f2 + 1)
    ]


def _coord(texto: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", texto.replace("$", ""))
    if not match:
        raise RuntimeError(f"Celda inválida {texto}")
    letras, fila = match.groups()
    numero = 0
    for letra in letras:
        numero = numero * 26 + (ord(letra) - 64)
    return numero, int(fila)


def _columna(numero: int) -> str:
    letras = ""
    while numero:
        numero, resto = divmod(numero - 1, 26)
        letras = chr(65 + resto) + letras
    return letras
