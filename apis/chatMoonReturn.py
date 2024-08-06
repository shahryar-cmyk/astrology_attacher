from flask import Blueprint, jsonify, request
import swisseph as swe

# Initialize the Flask Blueprint
chat_moon_return_calculation = Blueprint('chat_moon_return_calculation', __name__)


@chat_moon_return_calculation.route('/moon_return2', methods=['POST'])
def moon_return():
    data = request.get_json()

    birth_date = data['birth_date']  # Format: YYYY/MM/DD
    birth_time = data['birth_time']  # Format: HH:MM:SS
    current_year = data['current_year']

    # Convert date and time to Julian Day
    year, month, day = map(int, birth_date.split("/"))
    hour, minute, second = map(int, birth_time.split(":"))
    jd_birth = swe.julday(year, month, day, hour + minute / 60 + second / 3600)

    # Get the Sun position at birth
    sun_pos, ret = swe.calc_ut(jd_birth, swe.MOON)
    birth_sun_longitude = sun_pos[0]

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

    return jsonify({
        'moon_return_date': moon_return_date_str
    })