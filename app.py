from flask import Flask, request, jsonify
import subprocess
import re
import win32com.client
import pythoncom

app = Flask(__name__)

# For Getting the Planets Cordinates form the Swetest planet command
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
# Function to parse the output of the swetest planet command
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
                    minute = degree_match_min[0]
                    second = degree_match_min[1]
                    
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
        command = f"swetest -b{birth_date_day}.{birth_date_month}.{birth_date_year} -ut{ut_hour}:{ut_min}:{ut_sec} -p -house{lat_deg},{lon_deg},P -fPZ -roundsec"
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
        asteriod_12 = f"swetest -ps -xs511 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_13 = f"swetest -ps -xs704 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_14 = f"swetest -ps -xs107 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_15 = f"swetest -ps -xs65 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_16 = f"swetest -ps -xs121 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_17 = f"swetest -ps -xs10199 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_18 = f"swetest -ps -xs7 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_19 = f"swetest -ps -xs136108 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_20 = f"swetest -ps -xs136472 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_21 = f"swetest -ps -xs324 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_22 = f"swetest -ps -xs451 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_23 = f"swetest -ps -xs88 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_24 = f"swetest -ps -xs532 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_25 = f"swetest -ps -xs48 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_26 = f"swetest -ps -xs375 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_27 = f"swetest -ps -xs45 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_28 = f"swetest -ps -xs29 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_29 = f"swetest -ps -xs423 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_30 = f"swetest -ps -xs19 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_31 = f"swetest -ps -xs13 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_32 = f"swetest -ps -xs24 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_33 = f"swetest -ps -xs94 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_34 = f"swetest -ps -xs702 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_35 = f"swetest -ps -xs259 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_36 = f"swetest -ps -xs128 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_37 = f"swetest -ps -xs16 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_38 = f"swetest -ps -xs120 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_39 = f"swetest -ps -xs41 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_40 = f"swetest -ps -xs6 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_41 = f"swetest -ps -xs154 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_42 = f"swetest -ps -xs76 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_43 = f"swetest -ps -xs747 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_44 = f"swetest -ps -xs153 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_45 = f"swetest -ps -xs790 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_46 = f"swetest -ps -xs9 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_47 = f"swetest -ps -xs96 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_48 = f"swetest -ps -xs22 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_49 = f"swetest -ps -xs241 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_50 = f"swetest -ps -xs194 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_51 = f"swetest -ps -xs566 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_52 = f"swetest -ps -xs911 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_53 = f"swetest -ps -xs54 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_54 = f"swetest -ps -xs386 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_55 = f"swetest -ps -xs59 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_56 = f"swetest -ps -xs66652 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_57 = f"swetest -ps -xs47171 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_58 = f"swetest -ps -xs26308 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_59 = f"swetest -ps -xs65489 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # asteriod_60 = f"swetest -ps -xs88611 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_61 = f"swetest -ps -xs134860 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_62 = f"swetest -ps -xs130 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_63 = f"swetest -ps -xs42355 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_64 = f"swetest -ps -xs409 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_65 = f"swetest -ps -xs334 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_66 = f"swetest -ps -xs165 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_67 = f"swetest -ps -xs139 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_68 = f"swetest -ps -xs185 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_69 = f"swetest -ps -xs173 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_70 = f"swetest -ps -xs190 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        asteriod_71 = f"swetest -ps -xs536 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # asteriod_72 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # asteriod_73 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"
        # asteriod_74 = f"swetest -ps -xs5145 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"


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
        asteriod_11_result = subprocess.run(asteriod_11, shell=True, check=True, capture_output=True, text=True)
        asteriod_12_result = subprocess.run(asteriod_12, shell=True, check=True, capture_output=True, text=True)
        asteriod_13_result = subprocess.run(asteriod_13, shell=True, check=True, capture_output=True, text=True)
        asteriod_14_result = subprocess.run(asteriod_14, shell=True, check=True, capture_output=True, text=True)
        asteriod_15_result = subprocess.run(asteriod_15, shell=True, check=True, capture_output=True, text=True)
        asteriod_16_result = subprocess.run(asteriod_16, shell=True, check=True, capture_output=True, text=True)
        asteriod_17_result = subprocess.run(asteriod_17, shell=True, check=True, capture_output=True, text=True)
        asteriod_18_result = subprocess.run(asteriod_18, shell=True, check=True, capture_output=True, text=True)
        asteriod_19_result = subprocess.run(asteriod_19, shell=True, check=True, capture_output=True, text=True)
        asteriod_20_result = subprocess.run(asteriod_20, shell=True, check=True, capture_output=True, text=True)
        asteriod_21_result = subprocess.run(asteriod_21, shell=True, check=True, capture_output=True, text=True)
        asteriod_22_result = subprocess.run(asteriod_22, shell=True, check=True, capture_output=True, text=True)
        asteriod_23_result = subprocess.run(asteriod_23, shell=True, check=True, capture_output=True, text=True)
        asteriod_24_result = subprocess.run(asteriod_24, shell=True, check=True, capture_output=True, text=True)
        asteriod_25_result = subprocess.run(asteriod_25, shell=True, check=True, capture_output=True, text=True)
        asteriod_26_result = subprocess.run(asteriod_26, shell=True, check=True, capture_output=True, text=True)
        asteriod_27_result = subprocess.run(asteriod_27, shell=True, check=True, capture_output=True, text=True)
        asteriod_28_result = subprocess.run(asteriod_28, shell=True, check=True, capture_output=True, text=True)
        asteriod_29_result = subprocess.run(asteriod_29, shell=True, check=True, capture_output=True, text=True)
        asteriod_30_result = subprocess.run(asteriod_30, shell=True, check=True, capture_output=True, text=True)
        asteriod_31_result = subprocess.run(asteriod_31, shell=True, check=True, capture_output=True, text=True)
        asteriod_32_result = subprocess.run(asteriod_32, shell=True, check=True, capture_output=True, text=True)
        asteriod_33_result = subprocess.run(asteriod_33, shell=True, check=True, capture_output=True, text=True)
        asteriod_34_result = subprocess.run(asteriod_34, shell=True, check=True, capture_output=True, text=True)
        asteriod_35_result = subprocess.run(asteriod_35, shell=True, check=True, capture_output=True, text=True)
        asteriod_36_result = subprocess.run(asteriod_36, shell=True, check=True, capture_output=True, text=True)
        asteriod_37_result = subprocess.run(asteriod_37, shell=True, check=True, capture_output=True, text=True)
        asteriod_38_result = subprocess.run(asteriod_38, shell=True, check=True, capture_output=True, text=True)
        asteriod_39_result = subprocess.run(asteriod_39, shell=True, check=True, capture_output=True, text=True)
        asteriod_40_result = subprocess.run(asteriod_40, shell=True, check=True, capture_output=True, text=True)
        asteriod_41_result = subprocess.run(asteriod_41, shell=True, check=True, capture_output=True, text=True)
        asteriod_42_result = subprocess.run(asteriod_42, shell=True, check=True, capture_output=True, text=True)
        asteriod_43_result = subprocess.run(asteriod_43, shell=True, check=True, capture_output=True, text=True)
        asteriod_44_result = subprocess.run(asteriod_44, shell=True, check=True, capture_output=True, text=True)
        asteriod_45_result = subprocess.run(asteriod_45, shell=True, check=True, capture_output=True, text=True)
        asteriod_46_result = subprocess.run(asteriod_46, shell=True, check=True, capture_output=True, text=True)
        asteriod_47_result = subprocess.run(asteriod_47, shell=True, check=True, capture_output=True, text=True)
        asteriod_48_result = subprocess.run(asteriod_48, shell=True, check=True, capture_output=True, text=True)
        asteriod_49_result = subprocess.run(asteriod_49, shell=True, check=True, capture_output=True, text=True)
        asteriod_50_result = subprocess.run(asteriod_50, shell=True, check=True, capture_output=True, text=True)
        asteriod_51_result = subprocess.run(asteriod_51, shell=True, check=True, capture_output=True, text=True)
        asteriod_52_result = subprocess.run(asteriod_52, shell=True, check=True, capture_output=True, text=True)
        asteriod_53_result = subprocess.run(asteriod_53, shell=True, check=True, capture_output=True, text=True)
        asteriod_54_result = subprocess.run(asteriod_54, shell=True, check=True, capture_output=True, text=True)
        asteriod_55_result = subprocess.run(asteriod_55, shell=True, check=True, capture_output=True, text=True)
        asteriod_56_result = subprocess.run(asteriod_56, shell=True, check=True, capture_output=True, text=True)
        asteriod_57_result = subprocess.run(asteriod_57, shell=True, check=True, capture_output=True, text=True)
        asteriod_58_result = subprocess.run(asteriod_58, shell=True, check=True, capture_output=True, text=True)
        asteriod_59_result = subprocess.run(asteriod_59, shell=True, check=True, capture_output=True, text=True)
        # asteriod_60_result = subprocess.run(asteriod_60, shell=True, check=True, capture_output=True, text=True)
        asteriod_61_result = subprocess.run(asteriod_61, shell=True, check=True, capture_output=True, text=True)
        asteriod_62_result = subprocess.run(asteriod_62, shell=True, check=True, capture_output=True, text=True)
        asteriod_63_result = subprocess.run(asteriod_63, shell=True, check=True, capture_output=True, text=True)
        asteriod_64_result = subprocess.run(asteriod_64, shell=True, check=True, capture_output=True, text=True)
        asteriod_65_result = subprocess.run(asteriod_65, shell=True, check=True, capture_output=True, text=True)
        asteriod_66_result = subprocess.run(asteriod_66, shell=True, check=True, capture_output=True, text=True)
        asteriod_67_result = subprocess.run(asteriod_67, shell=True, check=True, capture_output=True, text=True)
        asteriod_68_result = subprocess.run(asteriod_68, shell=True, check=True, capture_output=True, text=True)
        asteriod_69_result = subprocess.run(asteriod_69, shell=True, check=True, capture_output=True, text=True)
        asteriod_70_result = subprocess.run(asteriod_70, shell=True, check=True, capture_output=True, text=True)
        asteriod_71_result = subprocess.run(asteriod_71, shell=True, check=True, capture_output=True, text=True)
        # asteriod_72_result = subprocess.run(asteriod_72, shell=True, check=True, capture_output=True, text=True)
        # asteriod_73_result = subprocess.run(asteriod_73, shell=True, check=True, capture_output=True, text=True)
        # asteriod_74_result = subprocess.run(asteriod_74, shell=True, check=True, capture_output=True, text=True)


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
        asteriod_11_output = asteriod_11_result.stdout
        asteriod_12_output = asteriod_12_result.stdout
        asteriod_13_output = asteriod_13_result.stdout
        asteriod_14_output = asteriod_14_result.stdout
        asteriod_15_output = asteriod_15_result.stdout
        asteriod_16_output = asteriod_16_result.stdout
        asteriod_17_output = asteriod_17_result.stdout
        asteriod_18_output = asteriod_18_result.stdout
        asteriod_19_output = asteriod_19_result.stdout
        asteriod_20_output = asteriod_20_result.stdout
        asteriod_21_output = asteriod_21_result.stdout
        asteriod_22_output = asteriod_22_result.stdout
        asteriod_23_output = asteriod_23_result.stdout
        asteriod_24_output = asteriod_24_result.stdout
        asteriod_25_output = asteriod_25_result.stdout
        asteriod_26_output = asteriod_26_result.stdout
        asteriod_27_output = asteriod_27_result.stdout
        asteriod_28_output = asteriod_28_result.stdout
        asteriod_29_output = asteriod_29_result.stdout
        asteriod_30_output = asteriod_30_result.stdout
        asteriod_31_output = asteriod_31_result.stdout
        asteriod_32_output = asteriod_32_result.stdout
        asteriod_33_output = asteriod_33_result.stdout
        asteriod_34_output = asteriod_34_result.stdout
        asteriod_35_output = asteriod_35_result.stdout
        asteriod_36_output = asteriod_36_result.stdout
        asteriod_37_output = asteriod_37_result.stdout
        asteriod_38_output = asteriod_38_result.stdout
        asteriod_39_output = asteriod_39_result.stdout
        asteriod_40_output = asteriod_40_result.stdout
        asteriod_41_output = asteriod_41_result.stdout
        asteriod_42_output = asteriod_42_result.stdout
        asteriod_43_output = asteriod_43_result.stdout
        asteriod_44_output = asteriod_44_result.stdout
        asteriod_45_output = asteriod_45_result.stdout
        asteriod_46_output = asteriod_46_result.stdout
        asteriod_47_output = asteriod_47_result.stdout
        asteriod_48_output = asteriod_48_result.stdout
        asteriod_49_output = asteriod_49_result.stdout
        asteriod_50_output = asteriod_50_result.stdout
        asteriod_51_output = asteriod_51_result.stdout
        asteriod_52_output = asteriod_52_result.stdout
        asteriod_53_output = asteriod_53_result.stdout
        asteriod_54_output = asteriod_54_result.stdout
        asteriod_55_output = asteriod_55_result.stdout
        asteriod_56_output = asteriod_56_result.stdout
        asteriod_57_output = asteriod_57_result.stdout
        asteriod_58_output = asteriod_58_result.stdout
        asteriod_59_output = asteriod_59_result.stdout
        # asteriod_60_output = asteriod_60_result.stdout
        asteriod_61_output = asteriod_61_result.stdout
        asteriod_62_output = asteriod_62_result.stdout
        asteriod_63_output = asteriod_63_result.stdout
        asteriod_64_output = asteriod_64_result.stdout
        asteriod_65_output = asteriod_65_result.stdout
        asteriod_66_output = asteriod_66_result.stdout
        asteriod_67_output = asteriod_67_result.stdout
        asteriod_68_output = asteriod_68_result.stdout
        asteriod_69_output = asteriod_69_result.stdout
        asteriod_70_output = asteriod_70_result.stdout
        asteriod_71_output = asteriod_71_result.stdout




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
        parsed_asteriod_11_output = parse_asteroid_output(asteriod_11_output)
        parsed_asteriod_12_output = parse_asteroid_output(asteriod_12_output)
        parsed_asteriod_13_output = parse_asteroid_output(asteriod_13_output)
        parsed_asteriod_14_output = parse_asteroid_output(asteriod_14_output)
        parsed_asteriod_15_output = parse_asteroid_output(asteriod_15_output)
        parsed_asteriod_16_output = parse_asteroid_output(asteriod_16_output)
        parsed_asteriod_17_output = parse_asteroid_output(asteriod_17_output)
        parsed_asteriod_18_output = parse_asteroid_output(asteriod_18_output)
        parsed_asteriod_19_output = parse_asteroid_output(asteriod_19_output)
        parsed_asteriod_20_output = parse_asteroid_output(asteriod_20_output)
        parsed_asteriod_21_output = parse_asteroid_output(asteriod_21_output)
        parsed_asteriod_22_output = parse_asteroid_output(asteriod_22_output)
        parsed_asteriod_23_output = parse_asteroid_output(asteriod_23_output)
        parsed_asteriod_24_output = parse_asteroid_output(asteriod_24_output)
        parsed_asteriod_25_output = parse_asteroid_output(asteriod_25_output)
        parsed_asteriod_26_output = parse_asteroid_output(asteriod_26_output)
        parsed_asteriod_27_output = parse_asteroid_output(asteriod_27_output)
        parsed_asteriod_28_output = parse_asteroid_output(asteriod_28_output)
        parsed_asteriod_29_output = parse_asteroid_output(asteriod_29_output)
        parsed_asteriod_30_output = parse_asteroid_output(asteriod_30_output)
        parsed_asteriod_31_output = parse_asteroid_output(asteriod_31_output)
        parsed_asteriod_32_output = parse_asteroid_output(asteriod_32_output)
        parsed_asteriod_33_output = parse_asteroid_output(asteriod_33_output)
        parsed_asteriod_34_output = parse_asteroid_output(asteriod_34_output)
        parsed_asteriod_35_output = parse_asteroid_output(asteriod_35_output)
        parsed_asteriod_36_output = parse_asteroid_output(asteriod_36_output)
        parsed_asteriod_37_output = parse_asteroid_output(asteriod_37_output)
        parsed_asteriod_38_output = parse_asteroid_output(asteriod_38_output)
        parsed_asteriod_39_output = parse_asteroid_output(asteriod_39_output)
        parsed_asteriod_40_output = parse_asteroid_output(asteriod_40_output)
        parsed_asteriod_41_output = parse_asteroid_output(asteriod_41_output)
        parsed_asteriod_42_output = parse_asteroid_output(asteriod_42_output)
        parsed_asteriod_43_output = parse_asteroid_output(asteriod_43_output)
        parsed_asteriod_44_output = parse_asteroid_output(asteriod_44_output)
        parsed_asteriod_45_output = parse_asteroid_output(asteriod_45_output)
        parsed_asteriod_46_output = parse_asteroid_output(asteriod_46_output)
        parsed_asteriod_47_output = parse_asteroid_output(asteriod_47_output)
        parsed_asteriod_48_output = parse_asteroid_output(asteriod_48_output)
        parsed_asteriod_49_output = parse_asteroid_output(asteriod_49_output)
        parsed_asteriod_50_output = parse_asteroid_output(asteriod_50_output)
        parsed_asteriod_51_output = parse_asteroid_output(asteriod_51_output)
        parsed_asteriod_52_output = parse_asteroid_output(asteriod_52_output)
        parsed_asteriod_53_output = parse_asteroid_output(asteriod_53_output)
        parsed_asteriod_54_output = parse_asteroid_output(asteriod_54_output)
        parsed_asteriod_55_output = parse_asteroid_output(asteriod_55_output)
        parsed_asteriod_56_output = parse_asteroid_output(asteriod_56_output)
        parsed_asteriod_57_output = parse_asteroid_output(asteriod_57_output)
        parsed_asteriod_58_output = parse_asteroid_output(asteriod_58_output)
        parsed_asteriod_59_output = parse_asteroid_output(asteriod_59_output)
        # parsed_asteriod_60_output = parse_asteroid_output(asteriod_60_output)
        parsed_asteriod_61_output = parse_asteroid_output(asteriod_61_output)
        parsed_asteriod_62_output = parse_asteroid_output(asteriod_62_output)
        parsed_asteriod_63_output = parse_asteroid_output(asteriod_63_output)
        parsed_asteriod_64_output = parse_asteroid_output(asteriod_64_output)
        parsed_asteriod_65_output = parse_asteroid_output(asteriod_65_output)
        parsed_asteriod_66_output = parse_asteroid_output(asteriod_66_output)
        parsed_asteriod_67_output = parse_asteroid_output(asteriod_67_output)
        parsed_asteriod_68_output = parse_asteroid_output(asteriod_68_output)
        parsed_asteriod_69_output = parse_asteroid_output(asteriod_69_output)
        parsed_asteriod_70_output = parse_asteroid_output(asteriod_70_output)
        parsed_asteriod_71_output = parse_asteroid_output(asteriod_71_output)


        # Return the parsed result as a JSON response
        return jsonify({"result": parsed_output,
                        "asteriod_Data": [parsed_asteriod_pholus_output,parsed_asteriod_1_output,parsed_asteriod_2_output,parsed_asteriod_3_output,parsed_asteriod_4_output,parsed_asteriod_5_output,parsed_asteriod_6_output,parsed_asteriod_7_output,parsed_asteriod_8_output,parsed_asteriod_9_output,parsed_asteriod_10_output,parsed_asteriod_11_output,parsed_asteriod_12_output,parsed_asteriod_13_output,parsed_asteriod_14_output,parsed_asteriod_15_output,parsed_asteriod_16_output,parsed_asteriod_17_output,parsed_asteriod_18_output,parsed_asteriod_19_output,parsed_asteriod_20_output,parsed_asteriod_21_output,parsed_asteriod_22_output,parsed_asteriod_23_output,parsed_asteriod_24_output,parsed_asteriod_25_output,parsed_asteriod_26_output,parsed_asteriod_27_output,parsed_asteriod_28_output,parsed_asteriod_29_output,parsed_asteriod_30_output,parsed_asteriod_31_output,parsed_asteriod_32_output,parsed_asteriod_33_output,parsed_asteriod_34_output,parsed_asteriod_35_output,parsed_asteriod_36_output,parsed_asteriod_37_output,parsed_asteriod_38_output,parsed_asteriod_39_output,parsed_asteriod_40_output,parsed_asteriod_41_output,parsed_asteriod_42_output,parsed_asteriod_43_output,parsed_asteriod_44_output,parsed_asteriod_45_output,parsed_asteriod_46_output,parsed_asteriod_47_output,parsed_asteriod_48_output,parsed_asteriod_49_output,parsed_asteriod_50_output,parsed_asteriod_51_output,parsed_asteriod_52_output,parsed_asteriod_53_output,parsed_asteriod_54_output,parsed_asteriod_55_output,parsed_asteriod_56_output,parsed_asteriod_57_output,parsed_asteriod_58_output,parsed_asteriod_59_output,parsed_asteriod_61_output,parsed_asteriod_62_output,parsed_asteriod_63_output,parsed_asteriod_64_output,parsed_asteriod_65_output,parsed_asteriod_66_output,parsed_asteriod_67_output,parsed_asteriod_68_output,parsed_asteriod_69_output,parsed_asteriod_70_output,parsed_asteriod_71_output]})

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
            pattern = r'\s{3,}'  # Pattern to split by 4 or more spaces
            match = re.split(pattern, lines[6])[1]
            name = re.split(pattern, lines[6])[0]
            degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", match)
            degree_match_sign = re.findall(r'[a-zA-Z]+', match)   
            degree_sign = degree_match_sign[0] if degree_match_sign else ""
            degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', match)
            degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
            degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
            degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
            # Only Teharonhiawako Left 
            # When the degree is not found with the first pattern, try the second pattern
            pattern1 = r'\s{2,}'  # Pattern to split by 3 or more spaces
            match1 = re.split(pattern1, lines[6])[1]
            degree_match1 = re.match(r"(\d{1,2})\s\w{2}\s.*", match1)
            # degree_match_sign1 = re.findall(r'[a-zA-Z]+', match1)   
            # degree_sign1 = degree_match_sign1[0] if degree_match_sign1 else ""
# import re

# # Define the string
# data = "Teharonhiawako 17 aq 19'59.3278"

# # Define the regex pattern
# pattern = r"(?P<name>[a-zA-Z]+)\s+(?P<degree>\d+)\s+(?P<sign>\w+)\s+(?P<min>\d+)'(?P<sec>[\d.]+)"

# # Use the pattern to search the data string
# match = re.search(pattern, data)

# # Extract the components if the pattern matches
# if match:
#     result = {
#         "name": match.group("name"),
#         "degree": match.group("degree"),
#         "sign": match.group("sign"),
#         "min": match.group("min"),
#         "sec": match.group("sec")
#     }
#     print(result)
# else:
#     print("No match found.")

            
            

            result[name] = {
                "name" : name,
                "position": {
                    "positionDegree": int(degree_match.group(1)) if degree_match else degree_match1.group(1),
                    "position_sign": degree_sign,
                    "position_min": degree_match_min[0],
                    "position_sec": degree_match_min[1] ,                    
                }
    
            }
        else:
            result["error"] = "Error parsing output: No lines in the output"
    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result[name]  # Always return a dictionary


def parse_house_output(output):
    lines = output.splitlines()  # Split by newline characters
    result = {}
    
    try:
        if len(lines) > 0:
            pattern = r'\s{3,}'  # Pattern to split by 3 or more spaces
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
                        "position_sec": degree_match_min[1] ,                 
                    }
        else:
            result["error"] = "Error parsing line: No lines in the output"
    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result  # Always return a dictionary

# Dummy user data (replace this with your actual data or database access)

# API to get a list of users
@app.route('/api/macro', methods=['GET'])
def run_excel_macro():
    pythoncom.CoInitialize()  # Initialize COM library
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True  # Set to True if you want Excel to be visible

        try:
            wb = xl.Workbooks.Open(r'C:\El Camino que Creas\Generador de Informes\Generador de Informes\Generador de Informes.xlsm')  # Path to your Excel file
            try:
                # Access the specific sheet by name
                sheet_name = 'CN y RS (o RL)'  # Replace with your sheet name
                sheet = wb.Sheets(sheet_name)

                # Modify data in the sheet
                # Example: Change cell A1 value to "Hello, World!"
                sheet.Range("S26").Value = "Hello, World!"

                print("Data modified successfully.")
                return jsonify({"message": "Data modified successfully."}), 200
            finally:                        
                wb.Close(SaveChanges=True)  # Save changes after running macro
        except Exception as e:
            print("Error opening workbook:", e)
            return jsonify({"error": str(e)}), 500
        finally:
            xl.Quit()
    except Exception as e:
        print("Error initializing Excel:", e)
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM library

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
