import requests
import json
 
url = "http://localhost:7071/api/contract_indexing_function"
headers = {
    "Content-Type": "application/json"
}
data = {"filename": "Billing_Doc0078.docx"}
 
response = requests.post(url, headers=headers, data=json.dumps(data))
 
print("Status Code:", response.status_code)
print("Response Body:", response.text)