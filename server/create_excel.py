import pymongo
from jsontoexcel import create_excel

client = pymongo.MongoClient('mongodb://localhost:27017')
db = client['funds_morningstar']
cursor = db.funds.find()
funds = list()
for document in cursor:
    funds.append(document)

create_excel(funds)
client.close()
