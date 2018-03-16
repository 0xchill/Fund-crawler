# server.py
from flask import Flask, render_template, request, redirect, Response
import random,json
import scrapy
from scrapy.crawler import CrawlerProcess
from quotes_spider import QuotesSpider


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

    process = CrawlerProcess()
    process.crawl(QuotesSpider)
    process.start()

    for item in data:
        # loop over every row
        result += str(item['make']) + '\n'

    return result

if __name__ == "__main__":
    app.run()
