url = "https://api.zigcore.com.br/integration/erp/lojas"

headers = {
    "Authorization": token
}

resp = requests.get(url, headers=headers)

print(resp.status_code)
print(resp.text)
