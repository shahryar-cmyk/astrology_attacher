from flask import Flask, jsonify

app = Flask(__name__)

@app.route('/second_endpoint', methods=['GET'])
def second_endpoint():
    # Define your second API logic here
    return jsonify({"message": "This is the second endpoint"})

