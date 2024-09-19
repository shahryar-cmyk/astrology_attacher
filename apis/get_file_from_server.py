from flask import Blueprint, send_file, jsonify, request
import os

get_report_file = Blueprint('getReportFile', __name__)

@get_report_file.route('/get_report_file', methods=['GET'])
def get_report_file_route():
    # Get the filename from the query parameters
    filename = request.args.get('filename')

    if not filename:
        return jsonify({"error": "Filename not provided"}), 400

    # # Define the base directory
    # base_directory = 'C:/El Camino que Creas/Generador de Informes'
    # file_path = os.path.join(base_directory, filename)

    # Check if the file exists before trying to send it
    if os.path.exists(filename):
        return send_file(filename, as_attachment=True)
    else:
        return jsonify({"error": "File not found"}), 404
