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
# Issue is in Uranus Seconds
    try:
        if len(lines) > 0:
            result["command"] = lines[0]
        if len(lines) > 1:
            result["date"] = lines[1]
        if len(lines) > 2:
            result["UT"] = lines[2]
        if len(lines) > 3:
            result["TT"] = lines[3]
        if len(lines) > 4:
            result["Epsilon"] = lines[4]
        if len(lines) > 5:
            result["Nutation"] = lines[5]
        if len(lines) > 6:
            result["Sun"] = parse_celestial_body(lines[6])
        if len(lines) > 7:
            result["Moon"] = parse_celestial_body(lines[7])
        if len(lines) > 8:
            result["Mercury"] = parse_celestial_body(lines[8])
        if len(lines) > 9:
            result["Venus"] = parse_celestial_body(lines[9])
        if len(lines) > 10:
            result["Mars"] = parse_celestial_body(lines[10])
        if len(lines) > 11:
            result["Jupiter"] = parse_celestial_body(lines[11])
        if len(lines) > 12:
            result["Saturn"] = parse_celestial_body(lines[12])
        if len(lines) > 13:
            result["Uranus"] = parse_celestial_body(lines[13])
        if len(lines) > 14:
            result["Neptune"] = parse_celestial_body(lines[14])
        if len(lines) > 15:
            result["Pluto"] = parse_celestial_body(lines[15])

    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result  # Always return a dictionary

def parse_celestial_body(line):
    try:
        # Split the line by spaces and filter out empty strings
        parts = [part.strip() for part in line.strip().split(" ") if part.strip()]
        if len(parts) >= 6:
            return {
                "name": parts[0],
                "position": parts[1],
                "longitude": parts[2],
                "latitude": parts[3],
                "speed": parse_speed(parts[4]),
                "distance": parse_distance(parts[5])
            }
    except Exception as e:
        return {"error": f"Error parsing celestial body line: {str(e)}"}

    return None


def parse_speed(speed_str):
    # Example speed strings: "-4° 7'28.47829196" or "-0° 0' 0.47457045"
    try:
        match = re.match(r"(-?\d+)°\s*(\d+)'(?:\s*(\d+\.\d+)|(\d+)\.(\d+))", speed_str)
        if match:
            degree, minutes, seconds_decimal, seconds_whole, seconds_fraction = match.groups()
            if seconds_decimal:
                seconds = float(seconds_decimal)
            else:
                seconds = float(f"{seconds_whole}.{seconds_fraction}")
            return {
                "degree": int(degree),
                "minutes": int(minutes),
                "seconds": seconds
            }
    except Exception as e:
        return {"error": f"Error parsing speed: {str(e)}"}

    return None

def parse_distance(distance_str):
    # Example distance strings: "17°27'55.38368914 11.07.1996 20:14:35 UT" or "-0°36'30.83018524"
    try:
        # Match pattern with optional date and time
        match = re.match(r"(-?\d+)°(\d+)'([\d\.]+)\s*(.*)", distance_str)
        if match:
            degree, minutes, seconds, date_time = match.groups()
            distance_info = {
                "degree": int(degree),
                "minutes": int(minutes),
                "seconds": float(seconds)
            }
            if date_time:  # Add date and time if present
                distance_info["date"] = date_time.strip()
            return distance_info
    except Exception as e:
        return {"error": f"Error parsing distance: {str(e)}"}

    return None

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
