# from flask import Blueprint, jsonify, request
# import swisseph as swe
# from datetime import datetime, timedelta
# import subprocess
# import re

# # Initialize the Flask Blueprint
# solar_return_calculation = Blueprint('solar_return_calculation', __name__)

# # File path for output
# output_file_path = 'solar_return_calculation_output.txt'

# def write_to_file(content):
#     with open(output_file_path, 'a') as file:
#         file.write(content + '\n')

# @solar_return_calculation.route('/calculate_solar_return', methods=['POST'])
# def calculate_solar_return():
#     try:
#         write_to_file("Start processing request")
        
#         # Get the data from the request
#         birth_date_year = int(request.json.get('birth_date_year'))
#         birth_date_month = int(request.json.get('birth_date_month'))
#         birth_date_day = int(request.json.get('birth_date_day'))
#         ut_hour = int(request.json.get('ut_hour'))
#         ut_min = int(request.json.get('ut_min'))
#         ut_sec = int(request.json.get('ut_sec'))
#         position_degree = int(request.json.get('positionDegree'))
#         position_min = float(request.json.get('position_min'))
#         position_sec = float(request.json.get('position_sec'))
#         position_sign = request.json.get('position_sign').lower()
#         current_year = int(request.json.get('current_year'))

#         write_to_file(f"Received input data: {request.json}")

#         # Zodiac sign degrees
#         zodiac_signs = {
#             "aries": 0,
#             "taurus": 30,
#             "gemini": 60,
#             "cancer": 90,
#             "leo": 120,
#             "virgo": 150,
#             "libra": 180,
#             "scorpio": 210,
#             "sagittarius": 240,
#             "capricorn": 270,
#             "aquarius": 300,
#             "pisces": 330
#         }

#         if position_sign not in zodiac_signs:
#             return jsonify({"error": "Invalid position sign provided."}), 400

#         # Convert position to total degrees
#         natal_sun_position = zodiac_signs[position_sign] + position_degree + (position_min / 60) + (position_sec / 3600)
#         write_to_file(f"Natal sun position: {natal_sun_position}")

#         swe.set_ephe_path('C:\\sweph\\ephe')

#         # Function to calculate the Julian Day Number
#         def julian_day(year, month, day, hour=0, minute=0, second=0):
#             return swe.julday(year, month, day, hour + minute / 60.0 + second / 3600.0)

#         # Your birth date and time
#         birth_date = datetime(birth_date_year, birth_date_month, birth_date_day, ut_hour, ut_min, ut_sec)
#         natal_jd = julian_day(birth_date.year, birth_date.month, birth_date.day,
#                               birth_date.hour, birth_date.minute, birth_date.second)
#         write_to_file(f"Natal Julian day: {natal_jd}")

#         # Get the current year
#         # current_year = datetime.now().year

#         # Start date for the search (15 days before birth date in current year)
#         start_date = datetime(current_year, birth_date.month, birth_date.day,
#                               birth_date.hour, birth_date.minute, birth_date.second) - timedelta(days=15)
#         start_jd = julian_day(start_date.year, start_date.month, start_date.day,
#                               start_date.hour, start_date.minute, start_date.second)
#         write_to_file(f"Start Julian day: {start_jd}")

#         # Find the two closest solar return dates within the 40-day window
#         closest_dates = []
#         for days in range(40):
#             for hour in range(24):
#                 jd = start_jd + days + (hour / 24.0)
#                 transiting_sun_position, _ = swe.calc(jd, swe.SUN)
#                 date = start_date + timedelta(days=days, hours=hour)
#                 positions = {
#                     "date": date.strftime('%Y-%m-%d %H:%M:%S'),
#                     "sun_position": transiting_sun_position[0]
#                 }
#                 diff = abs(transiting_sun_position[0] - natal_sun_position)
#                 if len(closest_dates) < 2:
#                     closest_dates.append((diff, positions))
#                     closest_dates.sort(key=lambda x: x[0])
#                 elif diff < closest_dates[-1][0]:
#                     closest_dates[-1] = (diff, positions)
#                     closest_dates.sort(key=lambda x: x[0])
#         write_to_file(f"Closest dates: {closest_dates}")

#         # Further refine the closest dates by minute and second
#         date_time_closest_date = datetime.strptime(closest_dates[0][1]["date"], '%Y-%m-%d %H:%M:%S')

#         # Run the command for minutes from swetest
#         command_for_min = (f"swetest -b{date_time_closest_date.day}.{date_time_closest_date.month}.2024 "
#                            f"-p0 -fTPZ -n60 -s1m -ut{date_time_closest_date.hour}:{date_time_closest_date.minute}:{date_time_closest_date.second} -ep")
#         result_output_min = subprocess.run(command_for_min, shell=True, check=True, capture_output=True, text=True)

#         # Log the command and output
#         write_to_file(f"Command for minutes: {command_for_min}")
#         write_to_file(f"Result output for minutes: {result_output_min.stdout}")

#         result_output_min_newLine = result_output_min.stdout.split('\n')

#         # Ensure there are enough lines
#         if len(result_output_min_newLine) < 6:
#             write_to_file(f"Error: Expected at least 6 lines in result_output_min_newLine but got {len(result_output_min_newLine)}")
#             return jsonify({"error": "Unexpected output from swetest command for minutes"}), 500

#         most_closest_date_min_dict = {}
#         for line in result_output_min_newLine[6:60]:
#             write_to_file(f"Processing line: {line}")
#             splitbySpace = re.split(r'Sun', line)

            
#             # Ensure the line split correctly
#             if len(splitbySpace) < 2:
#                 write_to_file(f"Error: Expected at least 2 parts after splitting line by 'Sun' but got {len(splitbySpace)}")
#                 continue
#             removeSpacing = splitbySpace[1].strip().replace(" ", "")
            
#             degreeSplit = re.split(r'\s\s+', removeSpacing)
#             write_to_file(f"Remove Spacing: {removeSpacing}")
            

#             pattern = r'[a-zA-Z]'
#             splitByAlphaBets = re.split(pattern, degreeSplit[0])
#             write_to_file(f"Split by Alphabets: {splitByAlphaBets}")
            
#             # Ensure the alphabet split correctly
#             if len(splitByAlphaBets) < 3:
#                 write_to_file(f"Error: Expected at least 3 parts after splitting by alphabets but got {len(splitByAlphaBets)}")
#                 continue
            
#             degreeMatch = splitByAlphaBets[0]
#             minSecMatch = re.split(r"'", splitByAlphaBets[2])
#             write_to_file(f"Min Sec Match: {minSecMatch}")
            
#             if len(minSecMatch) < 2:
#                 write_to_file(f"Error: Expected at least 2 parts after splitting by single quote but got min {len(minSecMatch)}")
#                 continue
            
#             most_closest_date_min_dict[splitbySpace[0]] = minSecMatch[1]

#         most_closest_date_min_find = find_closest_key(most_closest_date_min_dict, position_sec)
#         write_to_file(f"Dictory Iterate: {most_closest_date_min_dict}")
#         write_to_file(f"Min Found: {most_closest_date_min_find}")
#         removedUTfromString = re.split(r'UT', most_closest_date_min_find)
#         removedUTfromString = removedUTfromString[0].rstrip()
#         most_closest_date_min = datetime.strptime(f"{removedUTfromString}", "%d.%m.%Y %H:%M:%S")

#         write_to_file(f"Most closest date min: {most_closest_date_min}")

#         # Run the command for seconds from swetest
#         command_for_sec = (f"swetest -b{date_time_closest_date.day}.{date_time_closest_date.month}.2024 "
#                            f"-p0 -fTPZ -n60 -s1s -ut{date_time_closest_date.hour}:{most_closest_date_min.minute}:1 -ep")
#         result_output_sec = subprocess.run(command_for_sec, shell=True, check=True, capture_output=True, text=True)

#         # Log the command and output
#         write_to_file(f"Command for seconds: {command_for_sec}")
#         write_to_file(f"Result output for seconds: {result_output_sec.stdout}")

#         result_output_sec_newLine = result_output_sec.stdout.split('\n')

#         # Ensure there are enough lines
#         if len(result_output_sec_newLine) < 6:
#             write_to_file(f"Error: Expected at least 6 lines in result_output_sec_newLine but got {len(result_output_sec_newLine)}")
#             return jsonify({"error": "Unexpected output from swetest command for seconds"}), 500

#         most_closest_date_sec_dict = {}
#         for line in result_output_sec_newLine[6:60]:
#             write_to_file(f"Processing line: {line}")
#             splitbySpace = re.split(r'Sun', line)
            
#             if len(splitbySpace) < 2:
#                 write_to_file(f"Error: Expected at least 2 parts after splitting line by 'Sun' but got {len(splitbySpace)}")
#                 continue
#             removeSpacing = splitbySpace[1].strip().replace(" ", "")
            
#             degreeSplit = re.split(r'\s\s+', removeSpacing)
            

            
#             pattern = r'[a-zA-Z]'
#             splitByAlphaBets = re.split(pattern, degreeSplit[0])
         

            
#             if len(splitByAlphaBets) < 3:
#                 write_to_file(f"Error: Expected at least 3 parts after splitting by alphabets but got {len(splitByAlphaBets)}")
#                 continue
            
#             degreeMatch = splitByAlphaBets[0]
#             minSecMatch = re.split(r"'", splitByAlphaBets[2])
            
#             if len(minSecMatch) < 2:
#                 write_to_file(f"Error: Expected at least 2 parts after splitting by single quote but got sec {minSecMatch}")
#                 continue
            
#             most_closest_date_sec_dict[splitbySpace[0]] = minSecMatch[1]

#         most_closest_date_sec_find = find_closest_key(most_closest_date_sec_dict, position_sec)
#         write_to_file(f"Most closest date sec find: {most_closest_date_sec_find}")
#         removedUTfromString = re.split(r'UT', most_closest_date_sec_find)
#         removedUTfromString = removedUTfromString[0].rstrip()
#         most_closest_date_sec = datetime.strptime(f"{removedUTfromString}", "%d.%m.%Y %H:%M:%S")

#         write_to_file(f"Most closest date sec: {most_closest_date_sec}")

#         return jsonify({
#             "closest_date": most_closest_date_sec.strftime('%Y-%m-%d %H:%M:%S')
#         })

#     except Exception as e:
#         write_to_file(f"Error occurred: {str(e)}")
#         return jsonify({"error": "An error occurred during the calculation."}), 500

# # Helper function to find the closest key in the dictionary
# def find_closest_key(d, target):
#     closest_key = min(d.keys(), key=lambda k: abs(float(d[k]) - target))
#     return closest_key
