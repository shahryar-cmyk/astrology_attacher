from flask import Flask, request, jsonify
import subprocess

app = Flask(__name__)

@app.route('/testCommand', methods=['POST'])
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
        command = f"swetest -b{birth_date_day:02d}.{birth_date_month:02d}.{birth_date_year} -fPZgSBDT -ut{ut_hour:02d}:{ut_min:02d}:{ut_sec:02d} -ep"

        # Execute the command using subprocess
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        output = result.stdout
        error = result.stderr

        if error:
            return jsonify({"error": error}), 500

        # Return the result as a JSON response
        return jsonify({"result": output})

    except ValueError as e:
        return jsonify({"error": f"Invalid input type: {str(e)}"}), 400
    except subprocess.CalledProcessError as e:
        return jsonify({"error": f"Error executing command: {e.stderr}"}), 500
    except Exception as e:
        return jsonify({"error": f"An unexpected error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
