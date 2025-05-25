import requests
import json

url = f"http://172.16.20.73/AMG_API/auto/automation/automationapi/"
headers = {
    "Content-Type": "application/json"
}

def FAI_OK(FAI_OK_NUMBER):
    payload = {
    "Content": {
        "FAIType": "1",
        "StandardSampleID": F"{FAI_OK_NUMBER}",
        "Result": "00",
        "Equipment": "QAN13059"
    },
    "FunctionName": "EqpFAI",
    "FunctionUID": None,
    "FunctionType": "S"
    }
    return payload

def FAI_NG(FAI_NG_NUMBER):
    payload = {
    "Content": {
        "FAIType": "1",
        "StandardSampleID": F"{FAI_NG_NUMBER}",
        "Result": "01",
        "Equipment": "QAN13059"
    },
    "FunctionName": "EqpFAI",
    "FunctionUID": None,
    "FunctionType": "S"
    }
    return payload

while True:
    print("請輸入FAI OK輸入1，NG輸入2")
    inputee = input()
    if inputee == "1":
        payload = FAI_OK(FAI_NG_NUMBER = "GY6M9205K")
    elif inputee == "2":
        payload = FAI_NG(FAI_NG_NUMBER = "GY6M9205N")
    else :
        continue
    response = requests.post(url, headers=headers, json=payload)

    print("Status code:", response.status_code)
    try:
        print("Response JSON:", response.json())
    except json.JSONDecodeError:
        print("Response text:", response.text)
