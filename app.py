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
        # parsed_output = parse_swetest_output(output)

        # Return the parsed result as a JSON response
        return jsonify({"result": output})

    except ValueError as e:
        return jsonify({"error": f"Invalid input type: {str(e)}"}), 400
    except subprocess.CalledProcessError as e:
        return jsonify({"error": f"Error executing command: {e.stderr}"}), 500
    except Exception as e:
        return jsonify({"error": f"An unexpected error occurred: {str(e)}"}), 500

def parse_swetest_output(output):
    lines = output.split('\n')
    result = {}

    try:
        # Extract general information
        command_line = lines[0].split(': ')
        result["command"] = command_line[1] if len(command_line) > 1 else "Unknown command"

        date_line = lines[1].split()
        if len(date_line) >= 6:
            result["date"] = date_line[2]
            result["gregorian"] = date_line[2]
            result["UT"] = date_line[3]
            result["version"] = date_line[5]
        else:
            raise ValueError("Date line is not in the expected format")

        ut_julian_line = lines[2].split()
        if len(ut_julian_line) >= 5:
            result["UT_Julian"] = ut_julian_line[1]
            result["delta_t"] = ut_julian_line[4]
        else:
            raise ValueError("UT Julian line is not in the expected format")

        tt_julian_line = lines[3].split()
        if len(tt_julian_line) >= 2:
            result["TT_Julian"] = tt_julian_line[1]
        else:
            raise ValueError("TT Julian line is not in the expected format")

        epsilon_line = lines[4].split()
        if len(epsilon_line) >= 4:
            result["Epsilon_t"] = epsilon_line[2]
            result["Epsilon_m"] = epsilon_line[3]
        else:
            raise ValueError("Epsilon line is not in the expected format")

        nutation_line = lines[5].split()
        if len(nutation_line) >= 4:
            result["Nutation_longitude"] = nutation_line[1]
            result["Nutation_obliquity"] = nutation_line[3]
        else:
            raise ValueError("Nutation line is not in the expected format")

    except IndexError as e:
        raise ValueError(f"Error parsing general information: {str(e)}. Line: {lines}")

    celestial_bodies = {}
    for line in lines[7:]:
        if line.strip():
            try:
                parts = re.split(r'\s{2,}', line)
                if len(parts) >= 6:
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
                else:
                    raise ValueError("Celestial body line is not in the expected format")
            except IndexError as e:
                raise ValueError(f"Error parsing celestial body information: {str(e)}. Line: {line}")

    result.update(celestial_bodies)
    return result

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
