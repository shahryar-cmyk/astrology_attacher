from flask import Blueprint, jsonify, request
import swisseph as swe
from datetime import datetime, timedelta
import subprocess
import re
import math
# Initialize the Flask Blueprint
getDataExcel = Blueprint('getDataExcel', __name__)

# File path for output
output_file_path = 'getDataExcel_output.txt'

def write_to_file(content):
    with open(output_file_path, 'a') as file:
        file.write(content + '\n')

@getDataExcel.route('/get_data_excel', methods=['POST'])
def get_data_excel():
    print("Data modified successfully.",dir(swe))
    data = request.get_json()

    birth_date = data['birth_date']  # Format: YYYY/MM/DD
    birth_time = data['birth_time']  # Format: HH:MM:SS
    current_year = data['current_year']

    # Convert date and time to Julian Day
    year, month, day = map(int, birth_date.split("/"))
    hour, minute, second = map(int, birth_time.split(":"))
    
    
    jd_birth = swe.julday(year, month, day, hour)


    # Given degrees
    lat_deg = 74.55
    lon_deg = 32.4333333

# Conversion to radians
    lat_rad = lat_deg * math.pi / 180
    lon_rad = lon_deg * math.pi / 180

    # 
    flags = swe.FLG_EQUATORIAL | swe.FLG_SWIEPH | swe.FLG_SPEED 

    # Get the Sun position at birth
    sun_pos, ret = swe.calc_ut(jd_birth, swe.SUN)
    # Moon Positon
    moon_pos, ret = swe.calc_ut(jd_birth, swe.MOON)
    
    birth_sun_longitude = moon_pos[0]

    # Estimate Julian Day for the moon return (close to the birthday)
    jd_estimate = swe.julday(current_year, month, day)

    # Find the exact time the Sun returns to the same longitude using swe_solcross_ut
    serr = ''
    jd_moon_return = swe.mooncross_ut(birth_sun_longitude, jd_estimate, 0)

    if jd_moon_return < jd_estimate:
        return jsonify({'error': serr}), 400

    # Convert Julian Day to calendar date and time
    moon_return_date = swe.revjul(jd_moon_return)
    moon_return_date_str = f"{moon_return_date[0]}/{moon_return_date[1]:02d}/{moon_return_date[2]:02d} {int(moon_return_date[3])}:{int((moon_return_date[3] % 1) * 60):02d}:{int(((moon_return_date[3] % 1) * 60 % 1) * 60):02d}"
    # Split into degree
    # splitting decimal degrees into (zod. sign,) deg, min, sec. 
    degreeSun = swe.split_deg(sun_pos[0],swe.SPLIT_DEG_KEEP_DEG)
    

    return jsonify({
        'Sun': degreeSun,   
        'Moon' : moon_pos[0]
    })