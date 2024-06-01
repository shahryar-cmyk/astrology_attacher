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
        error = result.stderr

        if error:
            return jsonify({"error": error}), 500

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
    lines = output.split('\n')
    result = {}
    
    # Extract general information
    result["command"] = lines[0].split(': ')[1]
    result["date"] = lines[1].split()[2]
    result["gregorian"] = lines[1].split()[2]
    result["UT"] = lines[1].split()[3]
    result["version"] = lines[1].split()[5]
    result["UT_Julian"] = lines[2].split()[1]
    result["delta_t"] = lines[2].split()[4]
    result["TT_Julian"] = lines[3].split()[1]
    result["Epsilon_t"] = lines[4].split()[2]
    result["Epsilon_m"] = lines[4].split()[3]
    result["Nutation_longitude"] = lines[5].split()[1]
    result["Nutation_obliquity"] = lines[5].split()[3]

    # Extract celestial body information
    celestial_bodies = {}
    for line in lines[7:]:
        if line.strip():
            parts = re.split(r'\s{2,}', line)
            body = parts[0].split()[0]
            details = {
                "position": parts[0].split(' ', 1)[1],
                "longitude": parts[1],
                "latitude": parts[2],
                "speed": parts[3],
                "distance": parts[4],
                "date": parts[5]
            }
            celestial_bodies[body] = details
    
    result.update(celestial_bodies)

    return result

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
