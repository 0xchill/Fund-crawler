# coding: utf8
import scrapy
import csv
import os 
dir_path = os.path.dirname(os.path.realpath(__file__))


class QuotesSpider(scrapy.Spider):
    name = "isin"

    def start_requests(self):
        print(dir_path)
        url = "http://www.morningstar.fr/fr/funds/SecuritySearchResults.aspx?type=ALL&search="

        with open(dir_path+"\codes.csv","r") as f:
            reader = csv.DictReader(f,delimiter=";",fieldnames=["isin","code"])
            for row in reader:
                if reader.line_num > 1:
                    req = scrapy.Request(url=url+row["isin"], callback=self.parse)
                    req.meta['isin'] = row['isin']
                    yield req

    def parse(self, response):
        isin = response.meta['isin']

        try:
            table = response.xpath('//*[@id="ctl00_MainContent_fundTable"]')
            link = table.css('td').xpath('//*[@class="msDataText searchLink"]/a').xpath('@href')[0].extract()
            code = link.split("id=")[1]
        except IndexError:
            code = 'None'

        return {isin : code}



