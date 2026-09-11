"""
Capa de acceso al Excel adaptada a Sistema_Etoile_2.xlsx (Tienda Virtual).
Sin control de stock físico.

Estructura:
- INVENTARIO: CODIGO, PRENDA, COLOR, TALLA, PRECIO (Filas desde la 4 hasta la 144+)
- HISTORIAL: Registro de ventas (Filas de datos entre la 5 y la 204).
"""
import os
import shutil
import threading
from datetime import datetime
import re
import openpyxl
from openpyxl.utils import column_index_from_string

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Nombre y ruta del nuevo archivo Excel
EXCEL_PATH = os.environ.get(
    "ETOILE_EXCEL_PATH", os.path.join(BASE_DIR, "data", "Sistema_Etoile_2.xlsx")
)
BACKUP_DIR = os.path.join(BASE_DIR, "data", "backups")

_lock = threading.Lock()

# Configuración de INVENTARIO (según Sistema_Etoile_2.xlsx)
INV_HEADER_ROW = 3
INV_FIRST_DATA_ROW = 4
INV_COLS = {
    "CODIGO": 1,
    "PRENDA": 2,
    "COLOR": 3,
    "TALLA": 4,
    "PRECIO": 5,
}

# Configuración de HISTORIAL
HIST_HEADER_ROW = 4
HIST_FIRST_DATA_ROW = 5
HIST_LAST_DATA_ROW = 204
HIST_COLS = {
    "FECHA": 1,
    "CODIGO": 2,
    "PRENDA": 3,
    "COLOR": 4,
    "TALLA": 5,
    "CANT": 6,
    "PRECIO": 7,
    "SUBTOTAL": 8,
    "CLIENTE": 9,
    "PEDIDO_ID": 10,
    "TELEFONO": 11,
    "METODO_PAGO": 12,
    "NUM_TRANSACCION": 13,
}

_REF_SIMPLE = re.compile(r"^=([A-Z]+)(\d+)$")


def _backup():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(BACKUP_DIR, f"Sistema_Etoile_2_{stamp}.xlsx")
    try:
        shutil.copy2(EXCEL_PATH, dest)
        backups = sorted(
            f for f in os.listdir(BACKUP_DIR) if f.startswith("Sistema_Etoile_2_")
        )
        for old in backups[:-30]:
            os.remove(os.path.join(BACKUP_DIR, old))
    except FileNotFoundError:
        pass


def _load(data_only=False):
    return openpyxl.load_workbook(EXCEL_PATH, data_only=data_only)


def _resolver_celda(ws, row, col, visitados=None):
    if visitados is None:
        visitados = set()
    val = ws.cell(row=row, column=col).value
    if not (isinstance(val, str) and val.startswith("=")):
        return val
    m = _REF_SIMPLE.match(val)
    if not m:
        return None
    ref_col = column_index_from_string(m.group(1))
    ref_row = int(m.group(2))
    if (ref_row, ref_col) in visitados:
        return None
    visitados.add((ref_row, ref_col))
    return _resolver_celda(ws, ref_row, ref_col, visitados)


def get_inventario_raw(wb=None):
    """Lee el catálogo de productos tal cual está en el nuevo Excel."""
    own = wb is None
    wb = wb or _load(data_only=True)
    ws = wb["INVENTARIO"]
    items = []
    for row in range(INV_FIRST_DATA_ROW, ws.max_row + 1):
        codigo = ws.cell(row=row, column=INV_COLS["CODIGO"]).value
        if not codigo:
            continue
        items.append(
            {
                "codigo": str(codigo).strip(),
                "prenda": _resolver_celda(ws, row, INV_COLS["PRENDA"]),
                "color": _resolver_celda(ws, row, INV_COLS["COLOR"]),
                "talla": _resolver_celda(ws, row, INV_COLS["TALLA"]),
                "precio": ws.cell(row=row, column=INV_COLS["PRECIO"]).value or 0,
                "_row": row,
            }
        )
    if own:
        wb.close()
    return items


def get_historial_raw(wb=None):
    """Lee el registro de ventas en HISTORIAL."""
    own = wb is None
    wb = wb or _load(data_only=True)
    ws = wb["HISTORIAL"]
    items = []
    for row in range(HIST_FIRST_DATA_ROW, HIST_LAST_DATA_ROW + 1):
        codigo = ws.cell(row=row, column=HIST_COLS["CODIGO"]).value
        if not codigo or (isinstance(codigo, str) and codigo.startswith("=")):
            continue
        items.append(
            {
                "fecha": ws.cell(row=row, column=HIST_COLS["FECHA"]).value,
                "codigo": str(codigo).strip(),
                "prenda": ws.cell(row=row, column=HIST_COLS["PRENDA"]).value,
                "color": ws.cell(row=row, column=HIST_COLS["COLOR"]).value,
                "talla": ws.cell(row=row, column=HIST_COLS["TALLA"]).value,
                "cant": ws.cell(row=row, column=HIST_COLS["CANT"]).value or 0,
                "precio": ws.cell(row=row, column=HIST_COLS["PRECIO"]).value or 0,
                "subtotal": ws.cell(row=row, column=HIST_COLS["SUBTOTAL"]).value or 0,
                "cliente": ws.cell(row=row, column=HIST_COLS["CLIENTE"]).value,
                "pedido_id": ws.cell(row=row, column=HIST_COLS["PEDIDO_ID"]).value,
                "telefono": ws.cell(row=row, column=HIST_COLS["TELEFONO"]).value,
                "metodo_pago": ws.cell(row=row, column=HIST_COLS["METODO_PAGO"]).value,
                "num_transaccion": ws.cell(
                    row=row, column=HIST_COLS["NUM_TRANSACCION"]
                ).value,
            }
        )
    if own:
        wb.close()
    return items


def get_inventario_con_disponible():
    """Calcula unidades vendidas por producto (sin limitar por stock)."""
    with _lock:
        wb = _load(data_only=True)
        inventario = get_inventario_raw(wb)
        historial = get_historial_raw(wb)
        wb.close()

    vendido_por_codigo = {}
    for h in historial:
        vendido_por_codigo[h["codigo"]] = vendido_por_codigo.get(h["codigo"], 0) + (
            h["cant"] or 0
        )

    for item in inventario:
        item["vendido"] = vendido_por_codigo.get(item["codigo"], 0)

    return inventario


def get_prendas_agrupadas(busqueda=None):
    """Agrupa las variantes por tipo de prenda."""
    inventario = get_inventario_con_disponible()
    if busqueda:
        b = busqueda.strip().lower()
        inventario = [i for i in inventario if b in (i["prenda"] or "").lower()]

    agrupado = {}
    for item in inventario:
        nombre = item["prenda"] or "(sin nombre)"
        agrupado.setdefault(nombre, []).append(item)

    resultado = []
    for nombre, variantes in sorted(agrupado.items()):
        colores = sorted(set(v["color"] for v in variantes if v["color"]))
        tallas = sorted(set(v["talla"] for v in variantes if v["talla"]))
        precio = variantes[0]["precio"] if variantes else 0
        resultado.append(
            {
                "prenda": nombre,
                "colores": colores,
                "tallas": tallas,
                "precio": precio,
                "variantes": variantes,
            }
        )
    return resultado


def _abreviar(texto, largo):
    limpio = "".join(ch for ch in (texto or "") if ch.isalnum())
    return limpio[:largo].upper() if limpio else "XX"


def sugerir_codigo(prenda, color, talla, wb=None):
    palabras = [p for p in (prenda or "").split() if p]
    if len(palabras) >= 2:
        prendaAbrev = _abreviar(palabras[0], 2) + _abreviar(palabras[1], 2)
    elif len(palabras) == 1:
        prendaAbrev = _abreviar(palabras[0], 4)
    else:
        prendaAbrev = "XXXX"

    colorAbrev = _abreviar(color, 2)
    tallaTxt = "".join(ch for ch in (talla or "").strip().upper() if ch.isalnum()) or "XX"

    base = f"{prendaAbrev}-{colorAbrev}-{tallaTxt}"

    own = wb is None
    wb = wb or _load()
    inventario = get_inventario_raw(wb)
    if own:
        wb.close()
    existentes = {i["codigo"] for i in inventario}

    if base not in existentes:
        return base

    n = 2
    while f"{base}-{n}" in existentes:
        n += 1
    return f"{base}-{n}"


def get_prendas_nombres():
    inventario = get_inventario_raw()
    nombres = sorted(
        set(
            i["prenda"]
            for i in inventario
            if i["prenda"] and not str(i["prenda"]).startswith("=")
        )
    )
    return nombres


def agregar_prenda(prenda, color, talla, precio, codigo=None):
    """Agrega una nueva fila al catálogo INVENTARIO sin campo stock."""
    prenda = (prenda or "").strip()
    color = (color or "").strip()
    talla = (talla or "").strip()

    if not prenda:
        raise ValueError("Falta el nombre de la prenda")
    if precio is None or float(precio) < 0:
        raise ValueError("Precio inválido")

    with _lock:
        wb = _load()
        ws = wb["INVENTARIO"]
        inventario = get_inventario_raw(wb)
        existentes = {i["codigo"] for i in inventario}

        if codigo:
            codigo = codigo.strip().upper()
            if codigo in existentes:
                wb.close()
                raise ValueError(f"El código {codigo} ya existe.")
        else:
            codigo = sugerir_codigo(prenda, color, talla, wb)

        # Buscar la primera fila libre
        fila = INV_FIRST_DATA_ROW
        for row in range(INV_FIRST_DATA_ROW, ws.max_row + 2):
            if not ws.cell(row=row, column=INV_COLS["CODIGO"]).value:
                fila = row
                break
        else:
            fila = ws.max_row + 1

        ws.cell(row=fila, column=INV_COLS["CODIGO"], value=codigo)
        ws.cell(row=fila, column=INV_COLS["PRENDA"], value=prenda)
        ws.cell(row=fila, column=INV_COLS["COLOR"], value=color)
        ws.cell(row=fila, column=INV_COLS["TALLA"], value=talla)
        ws.cell(row=fila, column=INV_COLS["PRECIO"], value=float(precio))

        _backup()
        wb.save(EXCEL_PATH)
        wb.close()

    return codigo


def siguiente_pedido_id(wb):
    ws = wb["HISTORIAL"]
    max_id = 0
    for row in range(HIST_FIRST_DATA_ROW, HIST_LAST_DATA_ROW + 1):
        pid = ws.cell(row=row, column=HIST_COLS["PEDIDO_ID"]).value
        if pid:
            try:
                num = int(str(pid).replace("PED-", ""))
                max_id = max(max_id, num)
            except ValueError:
                continue
    return f"PED-{max_id + 1:04d}"


def crear_pedido(cliente, telefono, items, metodo_pago, num_transaccion):
    """
    Crea el pedido y registra la venta en HISTORIAL.
    Ya no restringe ni bloquea compras por stock.
    """
    if not items:
        raise ValueError("El pedido no tiene prendas")
    if metodo_pago not in ("efectivo", "transferencia"):
        raise ValueError("Método de pago inválido")
    if metodo_pago == "transferencia" and not num_transaccion:
        raise ValueError("Falta el número de transacción")

    with _lock:
        wb = _load()
        ws_hist = wb["HISTORIAL"]

        pedido_id = siguiente_pedido_id(wb)
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M")

        # Buscar la primera fila libre antes de la fila de resumen (204)
        fila = None
        for row in range(HIST_FIRST_DATA_ROW, HIST_LAST_DATA_ROW + 1):
            valor = ws_hist.cell(row=row, column=HIST_COLS["CODIGO"]).value
            if not valor or (isinstance(valor, str) and valor.startswith("=")):
                fila = row
                break

        filas_libres = fila is not None and (HIST_LAST_DATA_ROW - fila + 1) >= len(items)

        if fila is None or not filas_libres:
            wb.close()
            raise ValueError(
                f"La hoja HISTORIAL se quedó sin filas libres (hasta la fila {HIST_LAST_DATA_ROW})."
            )

        for it in items:
            subtotal = round(it["precio"] * it["cantidad"], 2)
            ws_hist.cell(row=fila, column=HIST_COLS["FECHA"], value=fecha)
            ws_hist.cell(row=fila, column=HIST_COLS["CODIGO"], value=it["codigo"])
            ws_hist.cell(row=fila, column=HIST_COLS["PRENDA"], value=it.get("prenda"))
            ws_hist.cell(row=fila, column=HIST_COLS["COLOR"], value=it.get("color"))
            ws_hist.cell(row=fila, column=HIST_COLS["TALLA"], value=it.get("talla"))
            ws_hist.cell(row=fila, column=HIST_COLS["CANT"], value=it["cantidad"])
            ws_hist.cell(row=fila, column=HIST_COLS["PRECIO"], value=it["precio"])
            ws_hist.cell(row=fila, column=HIST_COLS["SUBTOTAL"], value=subtotal)
            ws_hist.cell(row=fila, column=HIST_COLS["CLIENTE"], value=cliente)
            ws_hist.cell(row=fila, column=HIST_COLS["PEDIDO_ID"], value=pedido_id)
            ws_hist.cell(row=fila, column=HIST_COLS["TELEFONO"], value=telefono)
            ws_hist.cell(row=fila, column=HIST_COLS["METODO_PAGO"], value=metodo_pago)
            ws_hist.cell(
                row=fila,
                column=HIST_COLS["NUM_TRANSACCION"],
                value=num_transaccion if metodo_pago == "transferencia" else "EFECTIVO",
            )
            fila += 1

        _backup()
        wb.save(EXCEL_PATH)
        wb.close()

    return pedido_id


def get_pedidos_agrupados(desde=None, hasta=None, cliente=None, codigo=None):
    """Agrupa HISTORIAL por pedido_id para mostrar en el panel de ventas."""
    historial = get_historial_raw()

    def pasa_filtros(h):
        if desde and (not h["fecha"] or str(h["fecha"])[:10] < desde):
            return False
        if hasta and (not h["fecha"] or str(h["fecha"])[:10] > hasta):
            return False
        if cliente and cliente.lower() not in (h["cliente"] or "").lower():
            return False
        if codigo and codigo.lower() not in (h["codigo"] or "").lower():
            return False
        return True

    historial = [h for h in historial if pasa_filtros(h)]

    pedidos = {}
    for h in historial:
        pid = h["pedido_id"] or f"SIN-ID-{h['fecha']}-{h['cliente']}"
        if pid not in pedidos:
            pedidos[pid] = {
                "pedido_id": pid,
                "fecha": h["fecha"],
                "cliente": h["cliente"],
                "telefono": h["telefono"],
                "metodo_pago": h["metodo_pago"],
                "num_transaccion": h["num_transaccion"],
                "items": [],
                "total": 0,
            }
        pedidos[pid]["items"].append(h)
        pedidos[pid]["total"] += h["subtotal"] or 0


    def eliminar_pedido(pedido_id):
    """
    Elimina del HISTORIAL todas las filas asociadas al pedido_id indicado.
    Reorganiza las filas para mantener el rango de 5 a 204 limpio y sin huecos,
    reponiendo las fórmulas por defecto en las celdas liberadas.
    """
    pedido_id = str(pedido_id).strip()
    if not pedido_id:
        raise ValueError("ID de pedido no especificado.")

    with _lock:
        wb = _load(data_only=False)
        ws = wb["HISTORIAL"]

        # 1. Leer todas las filas reales existentes
        filas_conservadas = []
        encontrado = False

        for row in range(HIST_FIRST_DATA_ROW, HIST_LAST_DATA_ROW + 1):
            cod = ws.cell(row=row, column=HIST_COLS["CODIGO"]).value
            # Si no hay código o es fórmula no evaluada, no es dato real
            if not cod or (isinstance(cod, str) and cod.startswith("=")):
                continue

            pid = ws.cell(row=row, column=HIST_COLS["PEDIDO_ID"]).value
            if str(pid or "").strip() == pedido_id:
                encontrado = True
            else:
                # Guardamos los valores de la fila que NO se elimina (columnas 1 a 13)
                valores_fila = [ws.cell(row=row, column=c).value for c in range(1, 14)]
                filas_conservadas.append(valores_fila)

        if not encontrado:
            wb.close()
            raise ValueError(f"No se encontró el pedido '{pedido_id}' para eliminar.")

        # 2. Reescribir las filas conservadas desde HIST_FIRST_DATA_ROW en adelante
        curr_row = HIST_FIRST_DATA_ROW
        for fila_val in filas_conservadas:
            for col_idx, val in enumerate(fila_val, start=1):
                ws.cell(row=curr_row, column=col_idx, value=val)
            curr_row += 1

        # 3. Limpiar y restaurar fórmulas en las filas que quedaron vacías
        for row in range(curr_row, HIST_LAST_DATA_ROW + 1):
            ws.cell(row=row, column=HIST_COLS["FECHA"], value=None)
            ws.cell(row=row, column=HIST_COLS["CODIGO"], value=None)
            ws.cell(
                row=row,
                column=HIST_COLS["PRENDA"],
                value=f'=IF(B{row}="","",IFERROR(INDEX(INVENTARIO!$B:$B,MATCH(B{row},INVENTARIO!$A:$A,0)),"No encontrado"))'
            )
            ws.cell(
                row=row,
                column=HIST_COLS["COLOR"],
                value=f'=IF(B{row}="","",IFERROR(INDEX(INVENTARIO!$C:$C,MATCH(B{row},INVENTARIO!$A:$A,0)),""))'
            )
            ws.cell(
                row=row,
                column=HIST_COLS["TALLA"],
                value=f'=IF(B{row}="","",IFERROR(INDEX(INVENTARIO!$D:$D,MATCH(B{row},INVENTARIO!$A:$A,0)),""))'
            )
            ws.cell(row=row, column=HIST_COLS["CANT"], value=None)
            ws.cell(
                row=row,
                column=HIST_COLS["PRECIO"],
                value=f'=IF(B{row}="","",IFERROR(INDEX(INVENTARIO!$E:$E,MATCH(B{row},INVENTARIO!$A:$A,0)),""))'
            )
            ws.cell(
                row=row,
                column=HIST_COLS["SUBTOTAL"],
                value=f'=IF(B{row}="","",G{row}*F{row})'
            )
            ws.cell(row=row, column=HIST_COLS["CLIENTE"], value=None)
            ws.cell(row=row, column=HIST_COLS["PEDIDO_ID"], value=None)
            ws.cell(row=row, column=HIST_COLS["TELEFONO"], value=None)
            ws.cell(row=row, column=HIST_COLS["METODO_PAGO"], value=None)
            ws.cell(row=row, column=HIST_COLS["NUM_TRANSACCION"], value=None)

        _backup()
        wb.save(EXCEL_PATH)
        wb.close()

    return True
    resultado = list(pedidos.values())
    resultado.sort(key=lambda p: p["fecha"] or "", reverse=True)
    return resultado
