from flask import Flask, request, send_file, jsonify
from table3 import generate_ppt  # make sure this is the correct import

app = Flask(__name__)

@app.route("/generate_ppt", methods=["GET"])
def generate_ppt_api():
    master_client_id = request.args.get("master_client_id")
    booking_month = request.args.get("booking_month")
    client_id = request.args.get("client_id")
    is_client_level = request.args.get("is_client_level", "false").lower() == "true"

    # Validate required parameters
    if not master_client_id or not booking_month:
        return jsonify({"error": "master_client_id and booking_month are required"}), 400

    # Convert master_client_id to int
    try:
        master_client_id = int(master_client_id)
    except ValueError:
        return jsonify({"error": "master_client_id must be an integer"}), 400

    try:
        ppt_file_path = generate_ppt(master_client_id, booking_month, client_id, is_client_level)
        if ppt_file_path is None:
            return jsonify({"error": "Report generation failed. Possibly due to missing data."}), 500

        return send_file(ppt_file_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)

