from flask import Blueprint, jsonify, request
import swisseph as swe

# Initialize the Flask Blueprint
chat_solar_return_calculation = Blueprint('chat_solar_return_calculation', __name__)


@chat_solar_return_calculation.route('/solar_return2', methods=['POST'])
def solar_return():
    data = request.get_json()

    birth_date = data['birth_date']  # Format: YYYY/MM/DD
    birth_time = data['birth_time']  # Format: HH:MM:SS
    current_year = data['current_year']

    # Convert date and time to Julian Day
    year, month, day = map(int, birth_date.split("/"))
    hour, minute, second = map(int, birth_time.split(":"))
    jd_birth = swe.julday(year, month, day, hour + minute / 60 + second / 3600)

    # Get the Sun position at birth
    sun_pos, ret = swe.calc_ut(jd_birth, swe.SUN)
    birth_sun_longitude = sun_pos[0]

    # Estimate Julian Day for the solar return (close to the birthday)
    jd_estimate = swe.julday(current_year, month, day)

    # Find the exact time the Sun returns to the same longitude using swe_solcross_ut
    serr = ''
    jd_solar_return = swe.solcross_ut(birth_sun_longitude, jd_estimate, 0)

    if jd_solar_return < jd_estimate:
        return jsonify({'error': serr}), 400

    # Convert Julian Day to calendar date and time
    solar_return_date = swe.revjul(jd_solar_return)
    solar_return_date_str = f"{solar_return_date[0]}/{solar_return_date[1]:02d}/{solar_return_date[2]:02d} {int(solar_return_date[3])}:{int((solar_return_date[3] % 1) * 60):02d}:{int(((solar_return_date[3] % 1) * 60 % 1) * 60):02d}"

    return jsonify({
        'solar_return_date': solar_return_date_str
    })