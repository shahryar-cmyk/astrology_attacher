from flask import Flask, request, jsonify
import subprocess
import re

app = Flask(__name__)

@app.route('/testCommand', methods=['POST'])
def execute_command():
    try:
        # Get the parameters from the request data and ensure they are integers
        birth_date_year = int(request.json.get('birth_date_year'))
        birth_date_month = int(request.json.get('birth_date_month'))
        birth_date_day = int(request.json.get('birth_date_day'))
        ut_hour = int(request.json.get('ut_hour'))
        ut_min = int(request.json.get('ut_min'))
        ut_sec = int(request.json.get('ut_sec'))

        # Construct the command with zero-padded values
        command = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -fPZgSBDT -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -ep"

        # Execute the command using subprocess
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        output = result.stdout

        # Parse the output
        parsed_output = parse_swetest_output(output)

        # Return the parsed result as a JSON response
        return jsonify({"result": parsed_output})

    except ValueError as e:
        return jsonify({"error": f"Invalid input type: {str(e)}"}), 400
    except subprocess.CalledProcessError as e:
        return jsonify({"error": f"Error executing command: {e.stderr}"}), 500
    except Exception as e:
        return jsonify({"error": f"An unexpected error occurred: {str(e)}"}), 500

def parse_swetest_output(output):
    lines = output.splitlines()  # Split by newline characters
    result = {}

    try:
        # Example parsing logic, adapt this to the actual output format of swetest
        if len(lines) > 0:
            result["command"] = lines[0]
        if len(lines) > 1:
            result["date"] = lines[1]
        if len(lines) > 2:
            result["UT"] = lines[2]
        if len(lines) > 3:
            result["Nutation"] = lines[3]
        if len(lines) > 4:
            result["Sun"] = lines[4]
        if len(lines) > 5:
            result["Moon"] = lines[5]
        if len(lines) > 6:
            result["Mercury"] = lines[6]
        if len(lines) > 7:
            result["Venus"] = lines[7]

    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result  # Always return a dictionary

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
