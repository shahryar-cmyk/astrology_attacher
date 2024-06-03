from flask import Flask, request, jsonify
import subprocess
import re

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

                # Extract the degree part of the position
                degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", position)
                degree_match_sign = re.findall(r'[a-zA-Z]+', position)
                degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', position)
                degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
                degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
                degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
               
                if degree_match:
                    degree = int(degree_match.group(1))
                    degree_sign = degree_match_sign[0] if degree_match_sign else ""
                    min_sec_split = degree_match_min[0].split("'") if len(degree_match_min) > 1 else ["", ""]
                    minute = min_sec_split[0]
                    second = min_sec_split[1] if len(min_sec_split) > 1 else ""
                    
                    result[planet_name] = {
                        "positionDegree": degree,
                        "position_sign": degree_sign,
                        "position_min": minute,
                        "position_sec": second,
                    }
                else:
                    result[planet_name] = {"error": f"Error parsing degree from position: {position}"}
            else:
                result["error"] = f"Error parsing line: {line}"

    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result  # Always return a dictionary

@app.route('/house_endpoint', methods=['POST'])
def house_endpoint():
    try:
        # Get the parameters from the request data and ensure they are integers
        birth_date_year = int(request.json.get('birth_date_year'))
        birth_date_month = int(request.json.get('birth_date_month'))
        birth_date_day = int(request.json.get('birth_date_day'))
        ut_hour = int(request.json.get('ut_hour'))
        ut_min = int(request.json.get('ut_min'))
        ut_sec = int(request.json.get('ut_sec'))
        lat_deg = request.json.get('lat_deg')
        lon_deg = request.json.get('lon_deg')

        # Construct the command with zero-padded values
        command = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -p -house{lat_deg},{lon_deg},P -fPZ -roundsec"
        asteriod_pholus = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_1 = f"swetest -ps -xs136199 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_2 = f"swetest -ps -xs7066 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_3 = f"swetest -ps -xs28978 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_4 = f"swetest -ps -xs90482 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_5 = f"swetest -ps -xs50000 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_6 = f"swetest -ps -xs90377 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_7 = f"swetest -ps -xs10 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_8 = f"swetest -ps -xs87 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_9 = f"swetest -ps -xs31 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_10 = f"swetest -ps -xs15 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_11 = f"swetest -ps -xs624 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_12 = f"swetest -ps -xs52 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_13 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_14 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_15 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_16 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_17 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_18 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_19 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_20 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_21 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_22 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_23 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_24 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_25 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_26 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_27 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_28 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_29 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_30 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_31 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_32 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_33 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_34 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_35 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_36 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_37 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_38 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_39 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_40 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_41 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_42 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_43 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_44 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_45 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_46 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_47 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_48 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_49 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_50 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_51 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_52 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_53 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_54 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_55 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_56 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_57 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_58 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_59 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_60 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_61 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_62 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_63 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_64 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_65 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_66 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_67 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_68 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_69 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_70 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_71 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_72 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_73 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_74 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"


        # Execute the command using subprocess
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        asteriod_pholus_result = subprocess.run(asteriod_pholus, shell=True, check=True, capture_output=True, text=True)
        asteriod_1_result = subprocess.run(asteriod_1, shell=True, check=True, capture_output=True, text=True)
        asteriod_2_result = subprocess.run(asteriod_2, shell=True, check=True, capture_output=True, text=True)
        asteriod_3_result = subprocess.run(asteriod_3, shell=True, check=True, capture_output=True, text=True)
        asteriod_4_result = subprocess.run(asteriod_4, shell=True, check=True, capture_output=True, text=True)
        asteriod_5_result = subprocess.run(asteriod_5, shell=True, check=True, capture_output=True, text=True)
        asteriod_6_result = subprocess.run(asteriod_6, shell=True, check=True, capture_output=True, text=True)
        asteriod_7_result = subprocess.run(asteriod_7, shell=True, check=True, capture_output=True, text=True)
        asteriod_8_result = subprocess.run(asteriod_8, shell=True, check=True, capture_output=True, text=True)
        asteriod_9_result = subprocess.run(asteriod_9, shell=True, check=True, capture_output=True, text=True)
        asteriod_10_result = subprocess.run(asteriod_10, shell=True, check=True, capture_output=True, text=True)


        output = result.stdout
        asteriod_pholus_output = asteriod_pholus_result.stdout
        asteriod_1_output = asteriod_1_result.stdout
        asteriod_2_output = asteriod_2_result.stdout
        asteriod_3_output = asteriod_3_result.stdout
        asteriod_4_output = asteriod_4_result.stdout
        asteriod_5_output = asteriod_5_result.stdout
        asteriod_6_output = asteriod_6_result.stdout
        asteriod_7_output = asteriod_7_result.stdout
        asteriod_8_output = asteriod_8_result.stdout
        asteriod_9_output = asteriod_9_result.stdout
        asteriod_10_output = asteriod_10_result.stdout



        # Parse the output
        parsed_output = parse_house_output(output)
        # Parse the asteriod pholus output
        parsed_asteriod_pholus_output = parse_asteroid_output(asteriod_pholus_output)
        parsed_asteriod_1_output = parse_asteroid_output(asteriod_1_output)
        parsed_asteriod_2_output = parse_asteroid_output(asteriod_2_output)
        parsed_asteriod_3_output = parse_asteroid_output(asteriod_3_output)
        parsed_asteriod_4_output = parse_asteroid_output(asteriod_4_output)
        parsed_asteriod_5_output = parse_asteroid_output(asteriod_5_output)
        parsed_asteriod_6_output = parse_asteroid_output(asteriod_6_output)
        parsed_asteriod_7_output = parse_asteroid_output(asteriod_7_output)
        parsed_asteriod_8_output = parse_asteroid_output(asteriod_8_output)
        parsed_asteriod_9_output = parse_asteroid_output(asteriod_9_output)
        parsed_asteriod_10_output = parse_asteroid_output(asteriod_10_output)

        # Return the parsed result as a JSON response
        return jsonify({"result": parsed_output,
                        "asteriod_Data": [parsed_asteriod_pholus_output,parsed_asteriod_1_output,parsed_asteriod_2_output,parsed_asteriod_3_output,parsed_asteriod_4_output,parsed_asteriod_5_output,parsed_asteriod_6_output,parsed_asteriod_7_output,parsed_asteriod_8_output,parsed_asteriod_9_output,parsed_asteriod_10_output]})

    except ValueError as e:
        return jsonify({"error": f"Invalid input type: {str(e)}"}), 400
    except subprocess.CalledProcessError as e:
        return jsonify({"error": f"Error executing command: {e.stderr}"}), 500
    except Exception as e:
        return jsonify({"error": f"An unexpected error occurred: {str(e)}"}), 500


def parse_asteroid_output(asteroid_pholus_output):
    lines = asteroid_pholus_output.splitlines()  # Split by newline characters
    result = {}
    
    try:
        if len(lines) > 0:
            pattern = r'\s{3,}'  # Pattern to split by 3 or more spaces
            match = re.split(pattern, lines[6])[1]
            name = re.split(pattern, lines[6])[0]
            degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", match)
            degree_match_sign = re.findall(r'[a-zA-Z]+', match)   
            degree_sign = degree_match_sign[0] if degree_match_sign else ""
            degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', match)
            degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
            degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
            degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")

            result[name] = {
                    "positionDegree": int(degree_match.group(1)) if degree_match else None,
                    "position_sign": degree_sign,
                    "position_min": degree_match_min[0],
                    "position_sec": degree_match_min[1] if len(degree_match_min) > 1 else "",    
            }
        else:
            result["error"] = "Error parsing output: No lines in the output"
    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result  # Always return a dictionary


def parse_house_output(output):
    lines = output.splitlines()  # Split by newline characters
    result = {}
    
    try:
        if len(lines) > 0:
            pattern = r'\s{4,}'  # Pattern to split by 4 or more spaces
            result = {}
            for i in range(8, 14):  # Loop through lines 8 to 13 (houses 1 to 6)
                    match = re.split(pattern, lines[i])[1]
                    degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", match)
                    degree_match_sign = re.findall(r'[a-zA-Z]+', match)   
                    degree_sign = degree_match_sign[0] if degree_match_sign else ""
                    degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', match)
                    degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
                    degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
                    degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
                    # second = min_sec_split[1]                 
                    result[f"house{i - 7}"] = {
                        "positionDegree": int(degree_match.group(1)) if degree_match else None,
                        "position_sign": degree_sign,
                        "position_min": degree_match_min[0],
                        "position_sec": degree_match_min[1] if len(degree_match_min) > 1 else "",                 
                    }
        else:
            result["error"] = "Error parsing line: No lines in the output"
    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result  # Always return a dictionary


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
