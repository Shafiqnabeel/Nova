import os
from flask import Flask, jsonify, request, send_from_directory
import resend

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)

@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "nova.html")

@app.get("/nova.html")
def nova_page():
    return send_from_directory(BASE_DIR, "nova.html")

@app.get("/health")
def health():
    return jsonify({"ok": True, "configured": bool(os.getenv("RESEND_API_KEY", "").strip())})

@app.post("/send-email")
def send_email():
    try:
        api_key = os.getenv("RESEND_API_KEY", "").strip()
        if not api_key:
            return jsonify({"ok": False, "error": "RESEND_API_KEY is not configured on the server."}), 500

        resend.api_key = api_key
        payload = request.get_json(force=True) or {}
        recipient = (payload.get("to") or "").strip()
        subject = (payload.get("subject") or "Nova Shipment Update").strip()
        html = payload.get("html") or "<p>Nova shipment update.</p>"

        if "@" not in recipient:
            return jsonify({"ok": False, "error": "Invalid recipient email address."}), 400

        sender_name = os.getenv("NOVA_SENDER_NAME", "Nova | Digital Shipping Agent").strip()
        sender_email = os.getenv("NOVA_SENDER_EMAIL", "onboarding@resend.dev").strip()

        result = resend.Emails.send({
            "from": f"{sender_name} <{sender_email}>",
            "to": [recipient],
            "subject": subject,
            "html": html,
        })

        result_id = getattr(result, "id", None)
        if result_id is None and isinstance(result, dict):
            result_id = result.get("id")

        return jsonify({"ok": True, "id": result_id, "to": recipient})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port, debug=False)
