# coding: utf8
import scrapy
import json
import os
import sys
from scraper.items import FundItem,AssetItem,PerfItem,TopItem

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

    def get_allocation(self, response,fund):
        portfolio = response.xpath('//*[@id="overviewPortfolioAssetAllocationDiv"]/table')
        count = 0
        found = False
        actif = ""
        cat = 0
        # d_sub = dict()
        asset = AssetItem()
        fund['isin'] = response.meta['isin']
        fund['msCode'] = response.meta['code']
        for p in portfolio.css('td::text'):
            if p.extract() == "Actions":
                found = True
                cat = 1
            if found:
                if count == 0:
                    if cat != 3 :
                        actif = p.extract()
                    else:
                        actif = "liquidites"
                    count = count+1
                    cat = cat + 1
                elif count == 1:
                    # d_sub["Long"] = p.extract()
                    asset['long'] = p.extract()
                    count = count+1
                elif count == 2:
                    # d_sub["Short"] = p.extract()
                    asset['short'] = p.extract()
                    count = count+1
                elif count == 3:
                    # d_sub["Total"] = p.extract()
                    asset['total'] = p.extract()
                    count = 0
                    fund[actif.lower()] = dict(asset)
                    asset.clear()
                    # d_sub = {}

    def get_info(self,response,fund):
        #Nom du fond
        fund['nom'] = response.css('h1::text')[0].extract()

        # Valorisation du fond
        table = response.xpath('//*[@id="overviewQuickstatsDiv"]/table')
        fund['vl'] = table.css('tr')[1].css('td')[2].css('td::text')[0].extract()

    def get_perf(self,response,fund):
        # sub_d = {}
        perf = PerfItem()
        table = response.xpath('//*[@id="overviewTrailingReturnsDiv"]/table')

        perf["date"] = table.css('tr')[0].css('td')[1].css('td::text')[0].extract()
        # sub_d["Date"] = table.css('tr')[0].css('td')[1].css('td::text')[0].extract()
        keys = ['ytd','threeY','fiveY','tenY']
        for i in range(1,5):
            # period = table.css('tr')[i].css('td')[0].css('td::text')[0].extract()
            period=keys[i-1]
            value = table.css('tr')[i].css('td')[1].css('td::text')[0].extract()
            # sub_d[period] = value
            perf[period]=value

        fund['performance'] = dict(perf)

    def get_exposition(self,response,fund):
        regions = TopItem({'first':'ND','second':'ND','third':'ND','fourth':'ND','fifth':'ND'})
        secteurs = TopItem({'first':'ND','second':'ND','third':'ND','fourth':'ND','fifth':'ND'})

        keys=['first','second','third','fourth','fifth']

        try:
            table = response.xpath('//*[@id="overviewPortfolioTopRegionsDiv"]/table')
            for i in range(1,6):
                country = table.css('tr')[i].css('td')[0].css('td::text')[0].extract()
                weight = table.css('tr')[i].css('td')[1].css('td::text')[0].extract()
                regions[keys[i-1]]= country + " - " + weight.replace(",",'.') + " %"
                # sub_d[str(i)]= country + " - " + weight.replace(",",'.') + " %"

            # d['Top Regions'] = sub_d
            fund['regions']= dict(regions)
            regions.clear()
            # sub_d = {}
        except IndexError:
            # d['Top Regions'] = {"1":'-',"2":'-',"3":'-',"4":'-',"5":'-'}
            fund['regions'] = dict(regions)
            regions.clear()

        try:
            table = response.xpath('//*[@id="overviewPortfolioTopSectorsDiv"]/table')
            for i in range(1,6):
                sector = table.css('tr')[i].css('td')[1].css('td::text')[0].extract()
                weight = table.css('tr')[i].css('td')[2].css('td::text')[0].extract()
                secteurs[keys[i-1]]= sector + " - " + weight.replace(",",'.') + " %"
                # sub_d[str(i)]= sector + " - " + weight.replace(",",'.') + " %"

            # d['Top Secteurs'] = sub_d
            fund['secteurs'] = dict(secteurs)
            secteurs.clear()
        except IndexError:
            # d['Top Secteurs'] = {"1":'-',"2":'-',"3":'-',"4":'-',"5":'-'}
            fund['secteurs'] = dict(secteurs)
            secteurs.clear()


    def parse(self, response):
        fund = FundItem()
        self.get_info(response,fund)
        self.get_allocation(response,fund)
        self.get_perf(response,fund)
        self.get_exposition(response,fund)
        return fund
