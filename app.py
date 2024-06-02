from flask import Flask, request, jsonify
import subprocess
import re
# from apis.second_api import app as second_app 

app = Flask(__name__)

@app.route('/planets', methods=['POST'])
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
        planet_positions = lines[6:16]  # Adjust this slice according to the actual output format
        for line in planet_positions:
            # Extract planet name and position using regular expression
            match = re.match(r"(\w+)\s+(.+)", line)

            if match:
                planet_name = match.group(1)
                position = match.group(2).strip()
                position1 = match.group(2).strip()
                position2 = match.group(2).strip()



                # Extract the degree part of the position
                degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", position)
                degree_match_sign = re.findall(r'[a-zA-Z]+', position1)
                degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', position2)
                degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
                degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
                # Split the string by apostrophes
                degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
               

                

                if degree_match:
                    degree = int(degree_match.group(1))
                    degree_sign = degree_match_sign
                    result[planet_name] = {"positionDegree": degree,"position_sign":degree_sign[0], "position_min": degree_match_min[0],"position_sec": degree_match_min[1],}
                else:
                    result[planet_name] = {"error": f"Error parsing degree from position: {position}"}
            else:
                result["error"] = f"Error parsing line: {line}"

    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result  # Always return a dictionary

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
