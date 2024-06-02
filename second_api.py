from flask import Flask, jsonify

app = Flask(__name__)

@app.route('/second_endpoint', methods=['GET'])
def second_endpoint():
    # Define your second API logic here
    return jsonify({"message": "This is the second endpoint"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
