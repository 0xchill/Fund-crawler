# server.py
from flask import Flask, render_template, request, redirect, Response
import random,json
import pymongo


app = Flask(__name__, static_folder="../static/dist", template_folder="../static")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/hello")
def hello():
    return "Hello World!"

@app.route('/receiver', methods = ['POST'])
def worker():
    # read json + reply
    data = request.get_json()
    result = ''

    client = pymongo.MongoClient('mongodb://localhost:27017')
    db = client['funds_morningstar']
    cursor = db.funds.find()
    funds = list()
    for document in cursor:
        funds.append(document)

    print(funds[0])
    for item in data:
        # loop over every row
        result += str(item['make']) + '\n'

    return result

if __name__ == "__main__":
    app.run()
