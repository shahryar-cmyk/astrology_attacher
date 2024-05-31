# from flask import Flask

# app = Flask(__name__)

# @app.route('/')
# def hello_world():
#     return 'Hello World!'

# if __name__ == '__main__':
#     app.run(host='0.0.0.0', port=5000)

from flask import Flask, request, jsonify
import subprocess

app = Flask(__name__)

@app.route('/testCommand', methods=['POST'])
def execute_command():
    # Get the command from the request data
    command = request.json.get('command')

try:
    result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
    output = result.stdout
    error = result.stderr
    if error:
        return jsonify({"error": error}), 500
except subprocess.CalledProcessError as e:
    return jsonify({"error": f"Error executing command: {e.stderr}"}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
