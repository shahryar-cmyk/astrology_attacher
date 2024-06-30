from flask import Blueprint, jsonify, request
import subprocess
import re
import win32com.client
import pythoncom

change_excel_general = Blueprint('change_excel_general', __name__)

zodiac_signs = {
    'ar': 'Aries', 'ta': 'Tauro', 'ge': 'Géminis', 'cn': 'Cáncer',
    'le': 'Leo', 'vi': 'Virgo', 'li': 'Libra', 'sc': 'Escorpio',
    'sa': 'Sagitario', 'cp': 'Capricornio', 'aq': 'Acuario', 'pi': 'Piscis'
}

@change_excel_general.route('/change_excel_general', methods=['POST'])
def run_excel_macro_changeData():
    pythoncom.CoInitialize()
    try:
        data = request.json
        params = {
            'birth_date_year': int(data.get('birth_date_year')),
            'birth_date_month': int(data.get('birth_date_month')),
            'birth_date_day': int(data.get('birth_date_day')),
            'ut_hour': int(data.get('ut_hour')),
            'ut_min': int(data.get('ut_min')),
            'ut_sec': int(data.get('ut_sec')),
            'lat_deg': data.get('lat_deg'),
            'lon_deg': data.get('lon_deg')
        }

        commands = {
            'houses': f"swetest -b{params['birth_date_day']:02d}.{params['birth_date_month']:02d}.{params['birth_date_year']} "
                      f"-ut{params['ut_hour']:02d}:{params['ut_min']:02d}:{params['ut_sec']:02d} -p "
                      f"-house{params['lat_deg']},{params['lon_deg']},P -fPZ -roundsec",
            'planets': f"swetest -b{params['birth_date_day']:02d}.{params['birth_date_month']:02d}.{params['birth_date_year']} "
                       f"-fPZ -ut{params['ut_hour']:02d}:{params['ut_min']:02d}:{params['ut_sec']:02d} -ep",
            'quiron': f"swetest -ps -xs2060 -b{params['birth_date_day']:02d}.{params['birth_date_month']:02d}.{params['birth_date_year']} "
                      f"-ut{params['ut_hour']:02d}:{params['ut_min']:02d}:{params['ut_sec']:02d} -fPZ -roundsec"
        }

        results = {key: subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True).stdout.splitlines()
                   for key, cmd in commands.items()}

        def parse_output(output_lines, start, end, is_planet=False):
            parsed_data = {}
            pattern = r'\s{3,}'
            for i, line in enumerate(output_lines[start:end], start=1):
                try:
                    match = re.split(pattern, line)[1]
                    degree, sign, min_sec = re.match(r"(\d{1,2})\s(\w{2})\s(.*)", match).groups()
                    sign = zodiac_signs.get(sign.lower(), sign)
                    min, sec = re.split(r"[']", min_sec.replace(" ", ""))
                    parsed_data[f"Casa{i}"] = {
                        "positionDegree": int(degree),
                        "position_sign": sign,
                        "position_min": min,
                        "position_sec": sec.replace('"', '')
                    }
                    if is_planet:
                        parsed_data[f"Planet{i}"] = {"planet_name": line.split()[0], **parsed_data[f"Casa{i}"]}
                except (IndexError, AttributeError):
                    parsed_data[f"Casa{i}"] = {"error": "Error parsing line"}
            return parsed_data

        houses_data = parse_output(results['houses'], 8, 14)
        planets_data = parse_output(results['planets'], 6, 16, is_planet=True)
        quiron_data = parse_output(results['quiron'], 6, 7)

        file_path = r'C:\El Camino que Creas\Generador de Informes\Generador de Informes\Generador de Informes.xlsm'
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        wb = xl.Workbooks.Open(file_path)
        sheet = wb.Sheets['CN y RS (o RL)']

        for key, details in houses_data.items():
            if key.startswith("Casa"):
                row = int(key[-1]) + 4
                sheet.Range(f"E{row}").Value = details.get("positionDegree")
                sheet.Range(f"D{row}").Value = details.get("position_sign")
                sheet.Range(f"F{row}").Value = details.get("position_min")
                sheet.Range(f"G{row}").Value = details.get("position_sec")

        for idx, details in enumerate(planets_data.values(), start=1):
            if "error" not in details:
                row = idx + 28
                sheet.Range(f"R{row}").Value = details.get("planet_name")
                sheet.Range(f"U{row}").Value = details.get("positionDegree")
                sheet.Range(f"S{row}").Value = details.get("position_sign")
                sheet.Range(f"V{row}").Value = details.get("position_min")
                sheet.Range(f"W{row}").Value = details.get("position_sec")

        sheet.Range("R26").Value = quiron_data.get("name")
        sheet.Range("S26").Value = quiron_data.get("positionDegree")
        sheet.Range("T26").Value = quiron_data.get("position_sign")
        sheet.Range("U26").Value = quiron_data.get("position_min")

        wb.Close(SaveChanges=True)
        xl.Quit()
        return jsonify({"message": "Data modified successfully.", "result": houses_data, "result2": planets_data, "asteriods": [quiron_data]}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()

def parse_asteroid_output(asteroid_pholus_output):
    lines = asteroid_pholus_output.splitlines()
    result = {}
    try:
        pattern = r'\s{3,}'
        match = re.split(pattern, lines[6])[1]
        name = re.split(pattern, lines[6])[0]
        degree, sign, min_sec = re.match(r"(\d{1,2})\s(\w{2})\s(.*)", match).groups()
        sign = zodiac_signs.get(sign.lower(), sign)
        min, sec = re.split(r"[']", min_sec.replace(" ", ""))
        result[name] = {
            "name": name,
            "positionDegree": int(degree),
            "position_sign": sign,
            "position_min": min,
            "position_sec": sec
        }
    except (IndexError, AttributeError):
        result["error"] = "Error parsing output"
    return result[name]
