from flask import Blueprint, jsonify, request
import swisseph as swe
from datetime import datetime, timedelta
import subprocess
import re

# Initialize the Flask Blueprint
solar_return_calculation = Blueprint('solar_return_calculation', __name__)

# File path for output
output_file_path = 'solar_return_calculation_output.txt'

def write_to_file(content):
    with open(output_file_path, 'a') as file:
        file.write(content + '\n')

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
                jd = start_jd + days + (hour / 24.0)
                transiting_sun_position, _ = swe.calc(jd, swe.SUN)
                date = start_date + timedelta(days=days, hours=hour)
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
                
        # Further refine the closest dates by minute and second
        date_time_closest_date = datetime.strptime(closest_dates[0][1]["date"], '%Y-%m-%d %H:%M:%S')
        
        # Run the command for minutes from swetest
        command_for_min = (f"swetest -b{date_time_closest_date.day}.{date_time_closest_date.month}.2024 "
                           f"-p0 -fTPZ -n60 -s1m -ut{date_time_closest_date.hour}:{date_time_closest_date.minute}:{date_time_closest_date.second} -ep")
        result_output_min = subprocess.run(command_for_min, shell=True, check=True, capture_output=True, text=True)
        # write_to_file(f"result output for minutes: {result_output_min.stdout}")
        # Separate the output by new line
        result_output_min_newLine = result_output_min.stdout.split('\n')
        # Get the sixth line of the output
        most_closest_date_min = ''
        # Loop through sixth and forward to get the closest date
        for line in result_output_min_newLine[6:60]:
            # Seperate the line by space 5 or more
            # Split by spaces
            splitbySpace = re.split(r'Sun', line)
            
            # Seperate the degree and minutes and seconds 
            # Get the degree
            degreeSplit = re.split(r'\s\s+', splitbySpace[1])
            pattern = r'[a-zA-Z]'
            splitByAlphaBets = re.split(pattern, degreeSplit[1])
            # Degree
            degreeMatch = splitByAlphaBets[0]
            # Seperate by '
            minSecMatch = re.split(r"'", splitByAlphaBets[2])
            # Find the closest value of min given by the user
            if int(minSecMatch[0]) == position_min:
                # If the difference is less than 1 then write to the file
                if abs(float(minSecMatch[1]) - position_sec) < 1:

                    # write_to_file(f"Date: {splitbySpace[0]}  Degree: {degreeMatch} Minutes: {minSecMatch[0] } Seconds: {minSecMatch[1]}")
                    # Remove UT from the date
                    removedUTfromString = re.split(r'UT', splitbySpace[0])
                    # remove the space from the string at last 
                    removedUTfromString = removedUTfromString[0].rstrip()
                    most_closest_date_min = datetime.strptime(f"{removedUTfromString}", "%d.%m.%Y %H:%M:%S").minute
                    

        

        command_for_min = (f"swetest -b{date_time_closest_date.day}.{date_time_closest_date.month}.2024 "
                           f"-p0 -fTPZ -n60 -s1s -ut{date_time_closest_date.hour}:{most_closest_date_min}:1 -ep")
        result_output_sec = subprocess.run(command_for_min, shell=True, check=True, capture_output=True, text=True)
        result_output_sec_newLine = result_output_sec.stdout.split('\n')
        # Get the sixth line of the output
        most_closest_date_sec = ''
        for line in result_output_sec_newLine[6:60]:
            # Seperate the line by space 5 or more
            # Split by spaces
            splitbySpace = re.split(r'Sun', line)
            
            # Seperate the degree and minutes and seconds 
            # Get the degree
            degreeSplit = re.split(r'\s\s+', splitbySpace[1])
            pattern = r'[a-zA-Z]'
            splitByAlphaBets = re.split(pattern, degreeSplit[1])
            # Degree
            degreeMatch = splitByAlphaBets[0]
            # Seperate by '
            minSecMatch = re.split(r"'", splitByAlphaBets[2])
                # If the difference is less than 1 then write to the file
            if abs(float(minSecMatch[1]) - position_sec) < 1:

                    # write_to_file(f"Date: {splitbySpace[0]}  Degree: {degreeMatch} Minutes: {minSecMatch[0] } Seconds: {minSecMatch[1]}")
                    # Remove UT from the date
                    removedUTfromString = re.split(r'UT', splitbySpace[0])
                    # remove the space from the string at last 
                    removedUTfromString = removedUTfromString[0].rstrip()
                    most_closest_date_sec = datetime.strptime(f"{removedUTfromString}", "%d.%m.%Y %H:%M:%S").minute

            write_to_file(f"{result_output_min.stdout}")

                    
            


            


        

        # Run the command for seconds from swetest
        # command_for_sec = (f"swetest -b{date_time_closest_date.day}.{date_time_closest_date.month}.2024 "
        #                    f"-p0 -fTPZ -n60 -s1s -ut{date_time_closest_date.hour}:{date_time_closest_date.minute}:{date_time_closest_date.second} -ep")
        # result_output_sec = subprocess.run(command_for_sec, shell=True, check=True, capture_output=True, text=True)
        # write_to_file(f"result output for seconds: {result_output_sec.stdout}")

        response = {
            "closest_dates": closest_dates,
            "most_closest_date": {
                "date": closest_dates[0][1]["date"],
                "hour": date_time_closest_date.hour,
                "minute": most_closest_date_min
            },
            "total": natal_sun_position
        }

        return jsonify(response), 200

    except Exception as e:
        # Write the error to the file
        write_to_file(f"Error occurred: {str(e)}")
        return jsonify({"error": str(e)}), 500
