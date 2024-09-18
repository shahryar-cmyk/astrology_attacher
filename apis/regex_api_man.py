from flask import Blueprint, jsonify, request
import re
import os

# Define the Blueprint
regex_to_change_data = Blueprint('regex_to_change_data', __name__)

@regex_to_change_data.route('/regex_to_change_data', methods=['POST'])
def run_excel_macro_changeData():
    data = request.get_json()

    if not data or 'file_path' not in data:
        return jsonify({"error": "No file path provided"}), 400

    file_path = data['file_path']

    if not os.path.isfile(file_path):
        return jsonify({"error": "File does not exist"}), 400

    try:
        # Read the file content
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        # Define the regex pattern to find "PAK" and replace after it with spaces
        pattern = re.compile(r'^(.*?PAK = Pakistan)(.*)$')

        # Process each line individually
        processed_lines = []
        for line in lines:
            processed_lines.append(f'"{line.strip()}" =>"{line.strip()}",')

        # Write processed lines to a new file
        output_file_path = 'output.txt'
        with open(output_file_path, 'w', encoding='utf-8') as output_file:
            output_file.write('\n'.join(processed_lines) + '\n')

        return jsonify({"message": "File processed successfully", "output_file": output_file_path}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500
