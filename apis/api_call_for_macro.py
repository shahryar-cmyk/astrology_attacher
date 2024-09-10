from flask import Blueprint, request, jsonify
import requests
import subprocess  # Add this import for subprocess

debug_report_api_call = Blueprint('debugApiCall', __name__)

@debug_report_api_call.route('/debug_api_call', methods=['POST'])
def debug_report_api_call_route():
    # URL of the external API
    api_url = 'http://127.0.0.1:5000/get_report_ac_excel'

    try:
        # Parse request data
        data = request.get_json()
        file_name = data.get('fileName')
        macro_name = data.get('macroName')

        # Check if fileName and macroName are provided
        if not file_name or not macro_name:
            return jsonify({'error': 'fileName and macroName are required'}), 400
        
        # Prepare the data for the external API call
        data = {
            "fileName": file_name,
            "macroName": macro_name
        }
        
        # Check if another Excel process is running and terminate it
        try:
            subprocess.call(["taskkill", "/F", "/IM", "EXCEL.EXE"])
        except Exception as e:
            print("Error killing Excel process:", e)

        # Make the external API call
        response = requests.post(api_url, json=data)

        # Check if the request was successful
        if response.status_code == 200:
            # Process and return the response data
            return jsonify(response.json()), 200
        else:
            return jsonify({'error': f"Failed with status code {response.status_code}: {response.json()}"}), response.status_code

    except requests.exceptions.HTTPError as http_err:
        return jsonify({'error': str(http_err)}), 500
    except Exception as err:
        return jsonify({'error': str(err)}), 500
