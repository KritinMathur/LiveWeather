import json

with open('../bin/testlist.json',encoding = 'utf8') as f :
    data = json.load(f)

    for record in data['data']:
        print(record['id'],record['name'])

