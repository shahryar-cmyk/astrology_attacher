from flask import Blueprint, jsonify, request
import subprocess
import re
import win32com.client
import pythoncom

change_excel_general = Blueprint('change_excel_general', __name__)

# Mapeo de las abreviaturas de los signos del zodíaco a sus nombres completos
zodiac_signs = {
    'ar': 'Aries',
    'ta': 'Tauro',
    'ge': 'Géminis',
    'cn': 'Cáncer',
    'le': 'Leo',
    'vi': 'Virgo',
    'li': 'Libra',
    'sc': 'Escorpio',
    'sa': 'Sagitario',
    'cp': 'Capricornio',
    'aq': 'Acuario',
    'pi': 'Piscis'
    # Agrega otras abreviaturas si es necesario
}

@change_excel_general.route('/change_excel_general', methods=['POST'])
def run_excel_macro_changeData():
    pythoncom.CoInitialize()  # Initialize COM library
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

        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False  # Set to True if you want Excel to be visible

        # Construct the command with zero-padded values
        # For House Data From Cell D5 to D10
        command = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -p -house{lat_deg},{lon_deg},P -fPZ -roundsec"
        # For Planets Data From Cell D11 to D21 Which Includes True Node
        command2 = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -fPZ -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -ep"
        # For Quirón Command From Cell D22
        quiron_planet = f"swetest -ps -xs2060 -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -fPZ -roundsec"

        # Execute the command using subprocess
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        result2 = subprocess.run(command2, shell=True, check=True, capture_output=True, text=True)
        quiron_planet_result = subprocess.run(quiron_planet, shell=True, check=True, capture_output=True, text=True)

        output = result.stdout
        lines = output.splitlines()

        output2 = result2.stdout
        lines2 = output2.splitlines()

        quiron_output = quiron_planet_result.stdout
        quiron_parse_output= parse_asteroid_output(quiron_output)

        result_data = {}
        planets = []
        
        # Parse the output for houses
        if len(lines) > 0:
            pattern = r'\s{3,}'  # Pattern to split by 3 or more spaces
            for i in range(8, 14):  # Loop through lines 8 to 13 (houses 1 to 6)
                try:
                    match = re.split(pattern, lines[i])[1]
                    degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", match)
                    degree_match_sign = re.findall(r'[a-zA-Z]+', match)
                    degree_sign_abbr = degree_match_sign[0] if degree_match_sign else ""
                    degree_sign = zodiac_signs.get(degree_sign_abbr.lower(), degree_sign_abbr)
                    degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', match)
                    degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
                    degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
                    degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
                    result_data[f"Casa{i - 7}"] = {
                        "positionDegree": int(degree_match.group(1)) if degree_match else None,
                        "position_sign": degree_sign,
                        "position_min": degree_match_min[0],
                        "position_sec": degree_match_min[1].replace('"', ''),  # Remove double quotes from seconds
                    }
                except IndexError as e:
                    result_data["error"] = f"Error parsing output: {str(e)}"
                    break
        else:
            result_data["error"] = "Error parsing line: No lines in the output"
        
        # Parse the output for planets
        if len(lines2) > 0:
            planet_positions = lines2[6:16]
            planet_positions2 = lines2[17:18]

            for line in planet_positions:
                match = re.match(r"(\w+)\s+(.+)", line)
                if match:
                    planet_name = match.group(1)
                    position = match.group(2).strip()
                    degree_match = re.match(r"(\d{1,2})\s\w{2}\s.*", position)
                    degree_match_sign = re.findall(r'[a-zA-Z]+', position)
                    degree_sign_abbr = degree_match_sign[0] if degree_match_sign else ""
                    degree_sign = zodiac_signs.get(degree_sign_abbr.lower(), degree_sign_abbr)
                    degree_match_min_sec = re.sub(r'^.*?[a-zA-Z]', '', position)
                    degree_match_min_sec_again = re.sub(r'^.*?[a-zA-Z]', '', degree_match_min_sec)
                    degree_match_min_sec_again_spaces_removed = degree_match_min_sec_again.replace(" ", "")
                    degree_match_min = degree_match_min_sec_again_spaces_removed.split("'")
                    
                    if degree_match:
                        degree = int(degree_match.group(1))
                        minute = degree_match_min[0]
                        second = degree_match_min[1]
                        
                        planets.append({
                            "planet_name": planet_name,
                            "positionDegree": degree,
                            "position_sign": degree_sign,
                            "position_min": minute,
                            "position_sec": second,
                        })

            for line in planet_positions2:
                pattern = r"(True Node)\s+(\d+)\s+(\w+)\s+(\d+)\'([\d.]+)"
                match = re.match(pattern, line, re.IGNORECASE)
                if match:
                    planets.append({
                        "planet_name": match.group(1),
                        "positionDegree": match.group(2),
                        "position_sign": zodiac_signs.get(match.group(3), degree_sign_abbr),
                        "position_min": match.group(4),
                        "position_sec": match.group(5),
                    })
                else:
                    planets.append({"error": f"Error parsing line: {line}"})
        else:
            planets.append({"error": "Error parsing output for planets: No lines in the output"})

        # Open the workbook outside of the loop to avoid repeated opening and closing
        try:
            file_path = r'C:\El Camino que Creas\Generador de Informes\Generador de Informes\Generador de Informes.xlsm'
            wb = xl.Workbooks.Open(file_path)  # Path to your Excel file
            try:
                sheet_name = 'CN y RS (o RL)'  # Replace with your sheet name
                sheet = wb.Sheets(sheet_name)

                # Modify data in the sheet based on the result_data
                for casa, details in result_data.items():
                    if casa.startswith("Casa"):
                        degree_cell = f"E{int(casa[-1]) + 4}"  # Example cell for positionDegree
                        sign_cell = f"D{int(casa[-1]) + 4}"  # Example cell for position_sign
                        min_cell = f"F{int(casa[-1]) + 4}"  # Example cell for position_min
                        sec_cell = f"G{int(casa[-1]) + 4}"  # Example cell for position_sec

                        sheet.Range(degree_cell).Value = details["positionDegree"]
                        sheet.Range(sign_cell).Value = details["position_sign"]
                        sheet.Range(min_cell).Value = details["position_min"]
                        sheet.Range(sec_cell).Value = details["position_sec"]

                # Modify data in the sheet based on the planets
                planet_row_start = 20  # Example starting row for planet data
                for index, planet in enumerate(planets, start=1):
                    if "error" not in planet:
                        sheet.Range(f"R{index + 28}").Value = planet['planet_name']
                        sheet.Range(f"U{index + 28}").Value = planet['positionDegree']
                        sheet.Range(f"S{index + 28}").Value = planet['position_sign']
                        sheet.Range(f"V{index + 28}").Value = planet['position_min']
                        sheet.Range(f"W{index + 28}").Value = planet['position_sec']
                    else:
                        print(planet["error"])

                print("Data modified successfully.")
                return jsonify({"message": "Data modified successfully.", "result": result_data, "result2": planets, "asteriods": [quiron_parse_output] }), 200
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
                    "positionDegree": int(degree_match.group(1)) if degree_match else degree_match1.group(1),
                    "position_sign": degree_sign,
                    "position_min": degree_match_min[0],
                    "position_sec": degree_match_min[1] ,                    
                
    
            }
        else:
            result["error"] = "Error parsing output: No lines in the output"
    except IndexError as e:
        result["error"] = f"Error parsing output: {str(e)}"

    return result[name]  # Always return a dictionary