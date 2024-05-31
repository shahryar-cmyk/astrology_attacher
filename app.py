from flask import Flask, request, jsonify
import subprocess

app = Flask(__name__)

@app.route('/testCommand', methods=['POST'])
def execute_command():
    # Get the command from the request data
    command = request.json.get('command')

    # Execute the command using subprocess
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        output = result.stdout
        error = result.stderr
        if error:
            return jsonify({"error": error}), 500
    except subprocess.CalledProcessError as e:
        return jsonify({"error": f"Error executing command: {e.stderr}"}), 500

    # Return the result as a JSON response
    return jsonify({"result": output})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
