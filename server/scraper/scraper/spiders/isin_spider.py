# coding: utf8
import scrapy
import csv
import os
import pymongo
dir_path = os.path.dirname(os.path.realpath(__file__))


class IsinsSpider(scrapy.Spider):
    name = "isin"

    def start_requests(self):
        url = "http://www.morningstar.fr/fr/funds/SecuritySearchResults.aspx?type=ALL&search="
        self.mongo_uri='mongodb://localhost:27017'
        self.mongo_db='funds_morningstar'
        self.client = pymongo.MongoClient(self.mongo_uri)
        self.db = self.client[self.mongo_db]
        self.collection_name='isins'

        cursor = self.db.isins.find()

        # with open(dir_path+"\codes.csv","r") as f:
        #     reader = csv.DictReader(f,delimiter=";",fieldnames=["isin","code"])
        #     for row in reader:
        #         if reader.line_num > 1:
        #             req = scrapy.Request(url=url+row["isin"], callback=self.parse)
        #             req.meta['isin'] = row['isin']
        #             yield req
        for document in cursor:
            isin = document['isin']
            code = document['msCode']
            if code == 'None':
                req = scrapy.Request(url=url+isin, callback=self.parse)
                req.meta['isin'] = isin
                yield req
        self.client.close()


    def parse(self, response):
        isin = response.meta['isin']

        try:
            table = response.xpath('//*[@id="ctl00_MainContent_fundTable"]')
            link = table.css('td').xpath('//*[@class="msDataText searchLink"]/a').xpath('@href')[0].extract()
            code = link.split("id=")[1]
        except IndexError:
            code = 'None'

        self.db[self.collection_name].update_one({'isin':isin}, {'$set':{'msCode':code}})
