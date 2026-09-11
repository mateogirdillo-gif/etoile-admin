"""
Script de migración: agrega las columnas necesarias al Excel existente
sin dañar los datos ni las fórmulas que ya tienes.

Se corre UNA sola vez (o cada vez que quieras verificar que las columnas existan).
Uso:
    python migrate_excel.py
"""
import openpyxl
from openpyxl.styles import Font, PatternFill
import os

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "data", "Sistema_Etoile.xlsx")

HEADER_FILL = PatternFill(start_color="E8A1AC", end_color="E8A1AC", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def set_header(ws, col_idx, text):
    cell = ws.cell(row=ws._header_row if hasattr(ws, "_header_row") else None, column=col_idx)


def resolver_referencias_simples(ws):
    """
    Resuelve fórmulas del tipo '=B29' (una referencia directa a otra celda,
    usada aquí como atajo para no reescribir el mismo texto) y las reemplaza
    por su valor literal. Sigue la cadena hasta encontrar un valor real.
    Solo toca este patrón simple; no evalúa fórmulas complejas (SUMIFS, IF, etc.)
    porque esas viven en otras hojas (VENTAS/HISTORIAL) y no son datos maestros.
    """
    import re
    from openpyxl.utils import column_index_from_string

    patron = re.compile(r"^=([A-Z]+)(\d+)$")
    resueltas = 0

    def resolver(row, col, visitados):
        cell = ws.cell(row=row, column=col)
        val = cell.value
        if not (isinstance(val, str) and val.startswith("=")):
            return val
        m = patron.match(val)
        if not m:
            return val  # fórmula compleja, no la tocamos
        ref_col = column_index_from_string(m.group(1))
        ref_row = int(m.group(2))
        if (ref_row, ref_col) in visitados:
            return None  # referencia circular, evitar loop infinito
        visitados.add((ref_row, ref_col))
        return resolver(ref_row, ref_col, visitados)

    for row in range(INV_FIRST_DATA_ROW if False else 4, ws.max_row + 1):
        for col in range(1, 7):
            val = ws.cell(row=row, column=col).value
            if isinstance(val, str) and val.startswith("="):
                literal = resolver(row, col, {(row, col)})
                if literal is not None:
                    ws.cell(row=row, column=col, value=literal)
                    resueltas += 1
    return resueltas


def migrate():
    wb = openpyxl.load_workbook(EXCEL_PATH)

    # ---------- INVENTARIO: resolver fórmulas simples tipo '=B29' ----------
    ws_inv_pre = wb["INVENTARIO"]
    n = resolver_referencias_simples(ws_inv_pre)
    if n:
        print(f"✔ {n} celda(s) con fórmula simple resueltas a valor fijo en INVENTARIO")
    else:
        print("· INVENTARIO no tiene fórmulas simples pendientes")

    # ---------- INVENTARIO: agregar columna STOCK (F) ----------
    ws_inv = wb["INVENTARIO"]
    header_row_inv = 3  # fila donde está CODIGO, PRENDA, COLOR, TALLA, PRECIO
    headers_inv = [c.value for c in ws_inv[header_row_inv]]
    if "STOCK" not in headers_inv:
        col = headers_inv.index(None) + 1 if None in headers_inv else len(headers_inv) + 1
        # columna F (6) está vacía en el header actual, la usamos
        target_col = 6
        cell = ws_inv.cell(row=header_row_inv, column=target_col, value="STOCK")
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        # default 0 en filas con datos
        for row in range(header_row_inv + 1, ws_inv.max_row + 1):
            codigo = ws_inv.cell(row=row, column=1).value
            if codigo:
                existing = ws_inv.cell(row=row, column=target_col).value
                if existing is None:
                    ws_inv.cell(row=row, column=target_col, value=0)
        print("✔ Columna STOCK agregada a INVENTARIO (columna F)")
    else:
        print("· INVENTARIO ya tiene columna STOCK")

    # ---------- HISTORIAL: agregar PEDIDO_ID, TELEFONO, METODO_PAGO, NUM_TRANSACCION ----------
    ws_hist = wb["HISTORIAL"]
    header_row_hist = 4  # FECHA, CODIGO, PRENDA, COLOR, TALLA, CANT, PRECIO, SUBTOTAL, CLIENTE
    headers_hist = [c.value for c in ws_hist[header_row_hist]]

    new_cols = ["PEDIDO_ID", "TELEFONO", "METODO_PAGO", "NUM_TRANSACCION"]
    next_col = len([h for h in headers_hist if h is not None]) + 1

    for name in new_cols:
        if name not in headers_hist:
            cell = ws_hist.cell(row=header_row_hist, column=next_col, value=name)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            print(f"✔ Columna {name} agregada a HISTORIAL (columna {openpyxl.utils.get_column_letter(next_col)})")
            next_col += 1
        else:
            print(f"· HISTORIAL ya tiene columna {name}")

    wb.save(EXCEL_PATH)
    print("\nMigración completa. Archivo guardado:", EXCEL_PATH)


if __name__ == "__main__":
    migrate()
