from collections import Counter
from datetime import datetime
from urllib.parse import quote

from flask import Blueprint, current_app, flash, jsonify, redirect, render_template, request, session, url_for

from app.services.excel_service import excel_service
from app.utils.auth import login_required

admin_bp = Blueprint("admin", __name__, url_prefix="/admin")


def _to_str(value) -> str:
    return "" if value is None else str(value)


def _normalize_phone(raw_phone: str) -> str:
    digits = "".join(ch for ch in _to_str(raw_phone) if ch.isdigit())
    if digits.startswith("591"):
        digits = digits[3:]
    return digits


def _build_whatsapp_url(phone: str, codigo: str) -> str:
    normalized = _normalize_phone(phone)
    if not normalized:
        return ""
    message = f"Hola, te escribimos sobre tu pedido {codigo}"
    return f"https://wa.me/591{normalized}?text={quote(message)}"


def _parse_order_datetime(raw_value: str):
    raw = _to_str(raw_value).strip()
    if not raw:
        return None

    for date_format in ("%Y-%m-%d / %H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(raw, date_format)
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(raw)
    except ValueError:
        return None


def _compute_dashboard_metrics(orders):
    today = datetime.now(excel_service.tz_utc_minus_4).date()

    pending = sum(1 for o in orders if _to_str(o.get("estado")).lower() == "pendiente")
    delivered = sum(1 for o in orders if _to_str(o.get("estado")).lower() == "entregado")
    no_entregado = sum(1 for o in orders if _to_str(o.get("estado")).lower() == "no entregado")
    personal = sum(1 for o in orders if _to_str(o.get("tipo_entrega")).lower() == "entrega personal")
    envio = sum(1 for o in orders if _to_str(o.get("tipo_entrega")).lower() == "envio")

    pedidos_hoy = 0
    for order in orders:
        parsed = _parse_order_datetime(order.get("fecha_registro"))
        if parsed and parsed.date() == today:
            pedidos_hoy += 1

    by_status = Counter(_to_str(o.get("estado") or "Sin estado") for o in orders)

    return {
        "pending": pending,
        "delivered": delivered,
        "no_entregado": no_entregado,
        "today": pedidos_hoy,
        "personal": personal,
        "envio": envio,
        "by_status": dict(by_status),
    }


@admin_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        valid_user = current_app.config["ADMIN_USER"]
        valid_pass = current_app.config["ADMIN_PASSWORD"]

        if username == valid_user and password == valid_pass:
            session["admin_logged"] = True
            session["admin_user"] = username
            flash("Sesion iniciada correctamente.", "success")
            return redirect(url_for("admin.panel"))

        flash("Credenciales invalidas.", "danger")
        return redirect(url_for("admin.login"))

    return render_template("admin/login.html")


@admin_bp.route("/logout", methods=["POST"])
@login_required
def logout():
    session.clear()
    flash("Sesion cerrada correctamente.", "info")
    return redirect(url_for("admin.login"))


@admin_bp.route("", methods=["GET"])
@login_required
def panel():
    orders = excel_service.read_orders()
    recent_codes = excel_service.get_recent_codes(limit=25)
    metrics = _compute_dashboard_metrics(orders)
    options = excel_service.get_config_options()

    code_stats = Counter(_to_str(item.get("estado_codigo") or "Sin estado") for item in recent_codes)

    for order in orders:
        order["whatsapp_url"] = _build_whatsapp_url(
            phone=_to_str(order.get("telefono")),
            codigo=_to_str(order.get("codigo_producto")),
        )

    estados = [o["valor"] for o in options["estados"]]
    today_str = datetime.now(excel_service.tz_utc_minus_4).strftime("%Y-%m-%d")

    return render_template(
        "admin/panel.html",
        orders=orders,
        metrics=metrics,
        recent_codes=recent_codes,
        code_stats=dict(code_stats),
        options=options,
        estados=estados,
        today=today_str,
    )


@admin_bp.route("/api/orders", methods=["GET"])
@login_required
def api_orders():
    orders = excel_service.read_orders()
    recent_codes = excel_service.get_recent_codes(limit=25)
    metrics = _compute_dashboard_metrics(orders)
    code_stats = Counter(_to_str(item.get("estado_codigo") or "Sin estado") for item in recent_codes)

    for order in orders:
        order["whatsapp_url"] = _build_whatsapp_url(
            phone=_to_str(order.get("telefono")),
            codigo=_to_str(order.get("codigo_producto")),
        )

    return jsonify({
        "ok": True,
        "orders": orders,
        "recent_codes": recent_codes,
        "metrics": metrics,
        "code_stats": dict(code_stats),
        "estados": [item["valor"] for item in excel_service.get_config_options().get("estados", [])],
    })


@admin_bp.route("/codes/generate", methods=["POST"])
@login_required
def generate_code():
    generated_by = _to_str(session.get("admin_user") or "admin")
    code = excel_service.generate_next_code(generado_por=generated_by)
    return jsonify({"ok": True, "code": code.get("codigo")})


@admin_bp.route("/status", methods=["POST"])
@login_required
def update_status():
    data = request.get_json(silent=True) or {}
    order_id = str(data.get("order_id", "")).strip()
    new_status = str(data.get("status", "")).strip()

    valid_status = [item["valor"] for item in excel_service.get_config_options().get("estados", [])]

    if not order_id or not new_status:
        return jsonify({"ok": False, "message": "Datos incompletos"}), 400

    if new_status not in valid_status:
        return jsonify({"ok": False, "message": "Estado invalido"}), 400

    updated = excel_service.update_order_status(order_id, new_status)
    if not updated:
        return jsonify({"ok": False, "message": "No se encontro el pedido"}), 404

    return jsonify({"ok": True, "message": "Estado actualizado"})


@admin_bp.route("/config", methods=["GET", "POST"])
@login_required
def config_panel():
    if request.method == "POST":
        action = request.form.get("action", "").strip()
        sheet = request.form.get("sheet", "").strip()
        item_id = request.form.get("item_id", "").strip()
        value = request.form.get("value", "").strip()

        valid_sheets = set(excel_service.config_sheets.keys())
        if sheet not in valid_sheets:
            flash("Configuracion invalida.", "danger")
            return redirect(url_for("admin.config_panel"))

        try:
            if action == "add":
                excel_service.add_config_item(sheet, value)
                flash("Valor agregado correctamente.", "success")
            elif action == "edit":
                excel_service.update_config_item(sheet, item_id, value)
                flash("Valor actualizado correctamente.", "success")
            elif action == "toggle":
                active = request.form.get("active") == "true"
                excel_service.toggle_config_item(sheet, item_id, active)
                flash("Estado actualizado correctamente.", "info")
            elif action == "delete":
                excel_service.delete_config_item(sheet, item_id)
                flash("Valor eliminado correctamente.", "warning")
            else:
                flash("Accion no soportada.", "danger")
        except ValueError as err:
            flash(str(err), "danger")

        return redirect(url_for("admin.config_panel"))

    options = excel_service.get_config_options(include_inactive=True)
    return render_template("admin/config.html", options=options)
