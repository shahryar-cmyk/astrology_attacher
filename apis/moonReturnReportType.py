from flask import Blueprint, jsonify, request
import swisseph as swe
from datetime import datetime, timedelta
import logging

# Initialize the Flask Blueprint
solar_return_calculation = Blueprint('solar_return_calculation', __name__)

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

@solar_return_calculation.route('/calculate_solar_return', methods=['POST'])
def calculate_solar_return():
    try:
        # Get the data from the request
        birth_date_year = int(request.json.get('birth_date_year'))
        birth_date_month = int(request.json.get('birth_date_month'))
        birth_date_day = int(request.json.get('birth_date_day'))
        ut_hour = int(request.json.get('ut_hour'))
        ut_min = int(request.json.get('ut_min'))
        ut_sec = int(request.json.get('ut_sec'))
        position_degree = int(request.json.get('positionDegree'))
        position_min = float(request.json.get('position_min'))
        position_sec = float(request.json.get('position_sec'))
        position_sign = request.json.get('position_sign').lower()

        # Zodiac sign degrees
        zodiac_signs = {
            "aries": 0,
            "taurus": 30,
            "gemini": 60,
            "cancer": 90,
            "leo": 120,
            "virgo": 150,
            "libra": 180,
            "scorpio": 210,
            "sagittarius": 240,
            "capricorn": 270,
            "aquarius": 300,
            "pisces": 330
        }

        if position_sign not in zodiac_signs:
            return jsonify({"error": "Invalid position sign provided."}), 400

        # Convert position to total degrees
        natal_sun_position = zodiac_signs[position_sign] + position_degree + (position_min / 60) + (position_sec / 3600)

        swe.set_ephe_path('C:\\sweph\\ephe')

        # Function to calculate the Julian Day Number
        def julian_day(year, month, day, hour=0, minute=0, second=0):
            return swe.julday(year, month, day, hour + minute / 60.0 + second / 3600.0)

        # Your birth date and time
        birth_date = datetime(birth_date_year, birth_date_month, birth_date_day, ut_hour, ut_min, ut_sec)
        natal_jd = julian_day(birth_date.year, birth_date.month, birth_date.day,
                              birth_date.hour, birth_date.minute, birth_date.second)

        # Get the current year
        current_year = datetime.now().year

        # Start date for the search (15 days before birth date in current year)
        start_date = datetime(current_year, birth_date.month, birth_date.day,
                              birth_date.hour, birth_date.minute, birth_date.second) - timedelta(days=15)
        start_jd = julian_day(start_date.year, start_date.month, start_date.day,
                              start_date.hour, start_date.minute, start_date.second)

        # Find the two closest solar return dates within the 40-day window
        closest_dates = []
        for days in range(40):
            for hour in range(24):
                for minute in range(60):
                    for second in range(60):
                        jd = start_jd + days + (hour / 24.0) + (minute / 1440.0) + (second / 86400.0)
                        transiting_sun_position, _ = swe.calc(jd, swe.SUN)
                        date = start_date + timedelta(days=days, hours=hour, minutes=minute, seconds=second)
                        positions = {
                            "date": date.strftime('%Y-%m-%d %H:%M:%S'),
                            "sun_position": transiting_sun_position[0]
                        }
                        diff = abs(transiting_sun_position[0] - natal_sun_position)
                        if len(closest_dates) < 2:
                            closest_dates.append((diff, positions))
                            closest_dates.sort(key=lambda x: x[0])
                        elif diff < closest_dates[-1][0]:
                            closest_dates[-1] = (diff, positions)
                            closest_dates.sort(key=lambda x: x[0])

        # Get the most closest date
        most_closest_date = closest_dates[0][1] if closest_dates else None

        if most_closest_date:
            most_closest_datetime = datetime.strptime(most_closest_date["date"], '%Y-%m-%d %H:%M:%S')
            most_closest_date_details = {
                "year": most_closest_datetime.year,
                "month": most_closest_datetime.month,
                "day": most_closest_datetime.day,
                "hour": most_closest_datetime.hour,
                "minute": most_closest_datetime.minute,
                "second": most_closest_datetime.second
            }
        else:
            most_closest_date_details = None

        response = {
            "closest_dates": [date[1] for date in closest_dates],
            "most_closest_date": most_closest_date_details,
            "total": natal_sun_position
        }

        return jsonify(response), 200

    except Exception as e:
        logger.error(f"Error occurred: {str(e)}")
        return jsonify({"error": str(e)}), 500

# Example usage:
# from flask import Flask
# app = Flask()
# app.register_blueprint(solar_return_calculation, url_prefix='/api')
# if __name__ == '__main__':
#     app.run(debug=True)
