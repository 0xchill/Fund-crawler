# coding: utf8
import scrapy
import json
import os
dir_path = os.path.dirname(os.path.realpath(__file__))

class QuotesSpider(scrapy.Spider):
    name = "quotes"

    def start_requests(self):
        url = 'http://www.morningstar.fr/fr/funds/snapshot/snapshot.aspx?id='
        with open(dir_path + "\isin.json","r") as f:
            data = json.load(f)

        for pair in data:
            isin = list(pair)[0]
            code = pair[isin]
            if code != 'None':
                req = scrapy.Request(url=url+code, callback=self.parse)
                req.meta['isin'] = isin
                req.meta['code'] = code
                yield req

    def get_allocation(self, response,d):
        portfolio = response.xpath('//*[@id="overviewPortfolioAssetAllocationDiv"]/table')
        count = 0
        found = False
        actif = ""
        cat = 0
        d_sub = dict()
        d['ISIN'] = response.meta['isin']
        d['MS_CODE'] = response.meta['code']
        for p in portfolio.css('td::text'):
            if p.extract() == "Actions":
                found = True
                cat = 1
            if found:
                if count == 0:
                    if cat != 3 :
                        actif = p.extract()
                    else:
                        actif = "Liquidites"
                    count = count+1
                    cat = cat + 1
                elif count == 1:
                    d_sub["Long"] = p.extract()
                    count = count+1
                elif count == 2:
                    d_sub["Short"] = p.extract()
                    count = count+1
                elif count == 3:
                    d_sub["Total"] = p.extract()
                    count = 0
                    d[actif] = d_sub
                    d_sub = {}

    def get_info(self,response,d):
        #Nom du fond
        d["Nom"] = response.css('h1::text')[0].extract()

        # Valorisation du fond
        table = response.xpath('//*[@id="overviewQuickstatsDiv"]/table')
        d["VL"] = table.css('tr')[1].css('td')[2].css('td::text')[0].extract()

    def get_perf(self,response,d):
        sub_d = {}
        table = response.xpath('//*[@id="overviewTrailingReturnsDiv"]/table')

        sub_d["Date"] = table.css('tr')[0].css('td')[1].css('td::text')[0].extract()

        for i in range(1,5):
            period = table.css('tr')[i].css('td')[0].css('td::text')[0].extract()
            perf = table.css('tr')[i].css('td')[1].css('td::text')[0].extract()
            sub_d[period] = perf

        d['Performances'] = sub_d

    def get_exposition(self,response,d):
        sub_d = {}


        try:
            table = response.xpath('//*[@id="overviewPortfolioTopRegionsDiv"]/table')
            for i in range(1,6):
                country = table.css('tr')[i].css('td')[0].css('td::text')[0].extract()
                weight = table.css('tr')[i].css('td')[1].css('td::text')[0].extract()
                sub_d[str(i)]= country + " - " + weight.replace(",",'.') + " %"

            d['Top Regions'] = sub_d

            sub_d = {}
        except IndexError:
            d['Top Regions'] = {"1":'-',"2":'-',"3":'-',"4":'-',"5":'-'}

        try:
            table = response.xpath('//*[@id="overviewPortfolioTopSectorsDiv"]/table')
            for i in range(1,6):
                sector = table.css('tr')[i].css('td')[1].css('td::text')[0].extract()
                weight = table.css('tr')[i].css('td')[2].css('td::text')[0].extract()
                sub_d[str(i)]= sector + " - " + weight.replace(",",'.') + " %"

            d['Top Secteurs'] = sub_d
        except IndexError:
            d['Top Secteurs'] = {"1":'-',"2":'-',"3":'-',"4":'-',"5":'-'}


    def parse(self, response):
        # d = dict()
        # self.get_info(response,d)
        # self.get_allocation(response,d)
        # self.get_perf(response,d)
        # self.get_exposition(response,d)
        print("Existing settings: %s" % self.settings.attributes.keys())
        # return d
