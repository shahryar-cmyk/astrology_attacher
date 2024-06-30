from flask import Blueprint, jsonify

second_blueprint = Blueprint('second_blueprint', __name__)

@second_blueprint.route('/second_endpoint', methods=['GET'])
def second_endpoint():
    # Define your second API logic here
    return jsonify({"message": "This is the second endpoint"})