from datetime import datetime

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from app.services.excel_service import excel_service

user_bp = Blueprint("user", __name__)


def _sanitize_text(value: str) -> str:
    return " ".join((value or "").strip().split())


def _get_form_payload(form) -> dict:
    codigos_extra = form.getlist("codigos_extra")
    codigo_principal = excel_service.normalize_code(form.get("codigo_principal", ""))
    return {
        "codigo_principal": codigo_principal,
        "nombre": _sanitize_text(form.get("nombre", "")),
        "apellido": _sanitize_text(form.get("apellido", "")),
        "telefono": _sanitize_text(form.get("telefono", "")),
        "tipo_entrega": _sanitize_text(form.get("tipo_entrega", "")),
        "fecha_entrega": _sanitize_text(form.get("fecha_entrega", "")),
        "hora_entrega": _sanitize_text(form.get("hora_entrega", "")),
        "lugar_entrega": _sanitize_text(form.get("lugar_entrega", "")),
        "ciudad_destino": _sanitize_text(form.get("ciudad_destino", "")),
        "direccion_referencia": _sanitize_text(form.get("direccion_referencia", "")),
        "costo_total_pedido": _sanitize_text(form.get("costo_total_pedido", "")),
        "codigos_extra": [excel_service.normalize_code(code) for code in codigos_extra if _sanitize_text(code)],
    }


def _collect_codes(payload: dict) -> list:
    return [payload.get("codigo_principal", ""), *payload.get("codigos_extra", [])]


def _validate_payload(payload: dict) -> list:
    errors = []
    tipo = payload.get("tipo_entrega")

    required_common = [
        ("codigo_principal", "El codigo principal es obligatorio."),
        ("nombre", "El nombre es obligatorio."),
        ("telefono", "El numero de telefono es obligatorio."),
        ("costo_total_pedido", "El costo total del pedido es obligatorio."),
    ]

    for key, message in required_common:
        if not payload.get(key):
            errors.append(message)

    if tipo == "Entrega personal":
        if not payload.get("fecha_entrega"):
            errors.append("Debes seleccionar un dia de entrega.")
        if not payload.get("hora_entrega"):
            errors.append("Debes seleccionar una hora de entrega.")
        if not payload.get("lugar_entrega"):
            errors.append("Debes seleccionar un lugar de entrega.")
    elif tipo == "Envio":
        if not payload.get("apellido"):
            errors.append("El apellido es obligatorio para envios.")
        if not payload.get("ciudad_destino"):
            errors.append("Debes seleccionar una ciudad de destino.")
    else:
        errors.append("Selecciona un tipo de entrega valido.")

    if payload.get("codigo_principal") and not excel_service.is_valid_code_format(payload["codigo_principal"]):
        errors.append("El codigo debe tener el formato AA-00-A.")

    return errors


@user_bp.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        code = excel_service.normalize_code(request.form.get("codigo", ""))

        if not code:
            flash("Ingresa un codigo para continuar.", "warning")
            return redirect(url_for("user.home"))

        if not excel_service.is_valid_code_format(code):
            flash("El codigo debe tener el formato AA-00-A.", "warning")
            return redirect(url_for("user.home"))

        existing = excel_service.find_orders_by_code(code)
        if existing:
            return redirect(url_for("user.consulta", codigo=code))

        if not excel_service.is_code_available(code):
            flash("El codigo no existe, no esta disponible o ya fue usado.", "danger")
            return redirect(url_for("user.home"))

        return redirect(url_for("user.registro", codigo=code))

    return render_template("user/index.html")


@user_bp.route("/registro", methods=["GET"])
def registro():
    code = excel_service.normalize_code(request.args.get("codigo", ""))

    if code and not excel_service.is_code_available(code):
        existing = excel_service.find_orders_by_code(code)
        if existing:
            return redirect(url_for("user.consulta", codigo=code))

        flash("El codigo ingresado no esta disponible para registro.", "warning")
        return redirect(url_for("user.home"))

    config_options = excel_service.get_config_options()
    return render_template("user/register.html", code=code, options=config_options)


@user_bp.route("/registro/preview", methods=["POST"])
def registro_preview():
    payload = _get_form_payload(request.form)
    errors = _validate_payload(payload)

    if errors:
        for error in errors:
            flash(error, "danger")
        return redirect(url_for("user.registro", codigo=payload.get("codigo_principal")))

    if payload["codigo_principal"] in payload["codigos_extra"]:
        flash("No repitas el codigo principal en productos adicionales.", "warning")
        return redirect(url_for("user.registro", codigo=payload.get("codigo_principal")))

    for code in _collect_codes(payload):
        if not excel_service.is_code_available(code):
            flash(f"El codigo {code} no esta disponible o ya fue usado.", "danger")
            return redirect(url_for("user.registro", codigo=payload.get("codigo_principal")))

    session["preview_payload"] = payload

    resumen = {
        "cantidad_productos": 1 + len(payload.get("codigos_extra", [])),
        "codigos": [payload["codigo_principal"], *payload.get("codigos_extra", [])],
        "fecha_registro": datetime.now(excel_service.tz_utc_minus_4).strftime("%Y-%m-%d / %H:%M:%S"),
    }
    return render_template("user/preview.html", payload=payload, resumen=resumen)


@user_bp.route("/registro/confirm", methods=["POST"])
def registro_confirm():
    payload = session.get("preview_payload")
    if not payload:
        flash("No hay datos para confirmar. Completa el formulario nuevamente.", "warning")
        return redirect(url_for("user.home"))

    for code in _collect_codes(payload):
        if not excel_service.is_code_available(code):
            flash(f"El codigo {code} ya no esta disponible. Intenta con otro codigo.", "danger")
            return redirect(url_for("user.home"))

    group_id = excel_service.create_order_group(payload)

    for code in _collect_codes(payload):
        excel_service.mark_code_as_used(
            code=code,
            pedido_grupo_id=group_id,
            nombre=payload.get("nombre", ""),
            telefono=payload.get("telefono", ""),
        )

    session.pop("preview_payload", None)

    order = {
        "pedido_grupo_id": group_id,
        "codigo_principal": payload.get("codigo_principal"),
        "nombre": payload.get("nombre"),
        "telefono": payload.get("telefono"),
        "estado": "Pendiente",
        "tipo_entrega": payload.get("tipo_entrega"),
    }

    return render_template("user/success.html", order=order)


@user_bp.route("/consulta/<codigo>", methods=["GET"])
def consulta(codigo):
    code = _sanitize_text(codigo).upper()
    matches = excel_service.find_orders_by_code(code)

    if not matches:
        flash("No encontramos pedidos para ese codigo.", "warning")
        return redirect(url_for("user.home"))

    related = excel_service.get_related_orders(matches[0])
    principal = next((o for o in related if o.get("registro_principal") is True), related[0])

    return render_template("user/consulta.html", code=code, orders=related, principal=principal)
