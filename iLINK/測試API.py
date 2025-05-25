from flask import Flask, request, jsonify

app = Flask(__name__)
@app.route('/start_work', methods=['POST'])
def start_work():
    data = request.get_json()
    user_id = data.get("Content", {}).get("UserID", "")
    if user_id == "F1222":
        return jsonify({
            "FunctionName": "EquipmentAddUser",
            "FunctionUID": None,
            "ReturnCode": "00",
            "ReturnMessage": "OK",
            "FunctionType": "R",
            "Content": None
        })
    else:
        return jsonify({
            "FunctionName": "EquipmentAddUser",
            "FunctionUID": None,
            "ReturnCode": "01",
            "ReturnMessage": "無效的使用者",
            "FunctionType": "R",
            "Content": None
        })

@app.route('/start_program', methods=['POST'])
def start_program():
    return jsonify({
        "FunctionName": "EquipmentAddRecipe",
        "FunctionUID": None,
        "ReturnCode": "00",
        "ReturnMessage": "程式設定成功",
        "FunctionType": "R",
        "Content": None
    })

@app.route('/verify_operation', methods=['POST'])
def verify_operation():
    return jsonify({
        "FunctionName": "OperationVerify",
        "FunctionUID": None,
        "ReturnCode": "00",
        "ReturnMessage": "",
        "FunctionType": "R",
        "Content": {
            "DCSpec": None,
            "MMSSNCount": 0,
            "EDCParaCount": 1
        }
    })

@app.route('/execute_report', methods=['POST'])
def execute_report():
    return jsonify({
        "FunctionName": "OperationMove",
        "FunctionUID": None,
        "ReturnCode": "00",
        "ReturnMessage": "報工成功",
        "FunctionType": "R",
        "Content": None
    })

@app.route('/logout_user', methods=['POST'])
def logout_user():
    return jsonify({
        "FunctionName": "EquipmentRemoveUser",
        "FunctionUID": None,
        "ReturnCode": "00",
        "ReturnMessage": "下工成功",
        "FunctionType": "R",
        "Content": None
    })

@app.route('/clear_equipment', methods=['POST'])
def clear_equipment():
    return jsonify({
        "FunctionName": "EquipmentRemoveMLot",
        "FunctionUID": None,
        "ReturnCode": "00",
        "ReturnMessage": "清機完成",
        "FunctionType": "R",
        "Content": None
    })

if __name__ == '__main__':
    app.run(port=8000)
