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
        command = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -fPZ -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -ep"

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
        celestial_bodies = ["Sun", "Moon", "Mercury", "Venus", "Mars", "Jupiter", "Saturn", "Uranus", "Neptune", "Pluto"]
        for i, body in enumerate(celestial_bodies):
            if len(lines) > i + 6:
                result[body] = parse_celestial_body(body, lines[i + 6])
    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result

def parse_celestial_body(name, line):
    try:
        # Extract the components using regex
        match = re.match(r'([A-Za-z]+)\s+(\d+)\s+([a-z]{2})\s+(\d+)\'(\d+\.\d+)', line.strip())
        if match:
            position_degree, position_sign, position_minute, position_second = match.groups()[1:]
            return {
                "name": name,
                "positionDegree": int(position_degree),
                "positionSign": position_sign,
                "positionMinute": int(position_minute),
                "positionSecond": float(position_second)
            }
    except Exception as e:
        return {"error": f"Error parsing celestial body line: {str(e)}"}

    return None

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
