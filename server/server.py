# server.py
from flask import Flask, render_template, request, redirect, Response,send_file
import random,json
import pymongo
from jsontoexcel import create_excel
import time


app = Flask(__name__, static_folder="../static/dist", template_folder="../static")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/hello")
def hello():
    return "Hello World!"

@app.route('/receiver', methods = ['GET'])
def worker():

    client = pymongo.MongoClient('mongodb://localhost:27017')
    db = client['funds_morningstar']
    cursor = db.funds.find()
    funds = list()
    for document in cursor:
        funds.append(document)

    create_excel(funds)
    client.close()
    return ''

@app.route('/download', methods = ['GET'])
def download():
    xl = open('../static/Fonds.xlsx', "rb").read()

    return Response(
        xl,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-disposition":
                 "attachment; filename=fonds.xlsx"})
    # return send_file('../static/Fonds.xlsx', attachment_filename='Fonds.xlsx',as_attachment=True,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # print(funds[0])
    # for item in data:
    #     # loop over every row
    #     result += str(item['make']) + '\n'
    #
    # return result

if __name__ == "__main__":
    app.run(host= '192.168.1.10',port=5010)
