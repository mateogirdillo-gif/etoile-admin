import os
from flask import Flask, render_template, request, jsonify

import excel_service as xls

app = Flask(__name__)


# ---------------------------------------------------------------- páginas ----

@app.route("/")
def catalogo():
    return render_template("catalogo.html")


@app.route("/pedido")
def nuevo_pedido():
    return render_template("nuevo_pedido.html")


@app.route("/historial")
def historial():
    return render_template("historial.html")


# ------------------------------------------------------------------- API ----

@app.route("/api/prendas")
def api_prendas():
    busqueda = request.args.get("q", "").strip()
    try:
        data = xls.get_prendas_agrupadas(busqueda or None)
        return jsonify({"ok": True, "prendas": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/inventario")
def api_inventario():
    """Lista plana de variantes (para el buscador de agregar al pedido)."""
    try:
        data = xls.get_inventario_con_disponible()
        return jsonify({"ok": True, "items": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/stock", methods=["POST"])
def api_actualizar_stock():
    body = request.get_json(force=True)
    codigo = body.get("codigo")
    stock = body.get("stock")
    if codigo is None or stock is None:
        return jsonify({"ok": False, "error": "Faltan datos (codigo, stock)"}), 400
    try:
        xls.actualizar_stock(codigo, int(stock))
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


@app.route("/api/prendas_nombres")
def api_prendas_nombres():
    try:
        return jsonify({"ok": True, "nombres": xls.get_prendas_nombres()})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sugerir_codigo")
def api_sugerir_codigo():
    prenda = request.args.get("prenda", "")
    color = request.args.get("color", "")
    talla = request.args.get("talla", "")
    try:
        codigo = xls.sugerir_codigo(prenda, color, talla)
        return jsonify({"ok": True, "codigo": codigo})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/prenda", methods=["POST"])
def api_agregar_prenda():
    body = request.get_json(force=True)
    try:
        codigo = xls.agregar_prenda(
            prenda=body.get("prenda"),
            color=body.get("color"),
            talla=body.get("talla"),
            precio=body.get("precio"),
            stock=body.get("stock"),
            codigo=body.get("codigo") or None,
        )
        return jsonify({"ok": True, "codigo": codigo})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"Error inesperado: {e}"}), 500


@app.route("/api/pedido", methods=["POST"])
def api_crear_pedido():
    body = request.get_json(force=True)
    cliente = (body.get("cliente") or "").strip()
    telefono = (body.get("telefono") or "").strip()
    items = body.get("items") or []
    metodo_pago = body.get("metodo_pago")
    num_transaccion = (body.get("num_transaccion") or "").strip() or None

    if not cliente:
        return jsonify({"ok": False, "error": "Falta el nombre del cliente"}), 400
    if not items:
        return jsonify({"ok": False, "error": "El pedido no tiene prendas"}), 400

    try:
        pedido_id = xls.crear_pedido(
            cliente=cliente,
            telefono=telefono,
            items=items,
            metodo_pago=metodo_pago,
            num_transaccion=num_transaccion,
        )
        return jsonify({"ok": True, "pedido_id": pedido_id})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"Error inesperado: {e}"}), 500


@app.route("/api/historial")
def api_historial():
    desde = request.args.get("desde") or None
    hasta = request.args.get("hasta") or None
    cliente = request.args.get("cliente") or None
    codigo = request.args.get("codigo") or None
    try:
        pedidos = xls.get_pedidos_agrupados(desde, hasta, cliente, codigo)
        total_general = sum(p["total"] for p in pedidos)
        return jsonify({"ok": True, "pedidos": pedidos, "total_general": total_general})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/eliminar-pedido/<pedido_id>", methods=["POST"])
def ruta_eliminar_pedido(pedido_id):
    try:
        eliminar_pedido(pedido_id)
        return redirect(url_for("ver_pedidos"))  # o redireccionar al historial
    except Exception as e:
        return f"Error al eliminar: {str(e)}", 400

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
