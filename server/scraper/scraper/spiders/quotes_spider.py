# coding: utf8
import scrapy
import json
import os
import sys
import unidecode
import pymongo
import time
from scraper.items import FundItem,AssetItem,PerfItem,SecteurItem,GeoItem,TypeItem,DataItem

dir_path = os.path.dirname(os.path.realpath(__file__))

def format_string(s):
    return unidecode.unidecode(s.replace(' - ',' ').replace(' ','_').lower())

# Récupère les informations des actifs de type 'fund' sur Morningstar et met à jour
# la base de données Fair > isins en conséquence
class FundSpider(scrapy.Spider):
    name = "fund"

    def start_requests(self):
        url = 'http://www.morningstar.fr/fr/funds/snapshot/snapshot.aspx?id='

        self.mongo_uri='mongodb://localhost:27017'
        self.mongo_db='funds_morningstar'
        self.client = pymongo.MongoClient(self.mongo_uri)
        self.db = self.client[self.mongo_db]

        cursor = self.db.isins.find()

        for pair in cursor:
            isin=pair['isin']
            code=pair['msCode']
            if code != 'None':
                req = scrapy.Request(url=url+code, callback=self.parse)
                req.meta['isin'] = isin
                req.meta['code'] = code
                yield req
        self.client.close()

        # with open(dir_path + "\isin.json","r") as f:
        #     data = json.load(f)
            # for pair in data:
            #     isin = list(pair)[0]
            #     code = pair[isin]
            #     if code != 'None':
            #         req = scrapy.Request(url=url+code, callback=self.parse)
            #         req.meta['isin'] = isin
            #         req.meta['code'] = code
            #         yield req
    def get_info(self,response,fund):
        #Nom du fond
        fund['nom'] = response.css('h1::text')[0].extract()

        #Infos présentes dans la requète
        fund['isin'] = response.meta['isin']
        fund['msCode'] = response.meta['code']

    def get_allocation(self, response,data):
        portfolio = response.xpath('//*[@id="overviewPortfolioAssetAllocationDiv"]/table')
        count = 0
        found = False
        actif = ""
        cat = 0
        repartition = TypeItem()
        # d_sub = dict()
        asset = AssetItem()

        for p in portfolio.css('td::text'):
            if p.extract() == "Actions":
                found = True
                cat = 1
            if found:
                if count == 0:
                    if cat != 3 :
                        actif = p.extract().replace(",",'.')
                    else:
                        actif = "liquidites"
                    count = count+1
                    cat = cat + 1
                elif count == 1:
                    # d_sub["Long"] = p.extract()
                    asset['long'] = float(p.extract().replace(",",'.'))
                    count = count+1
                elif count == 2:
                    # d_sub["Short"] = p.extract()
                    asset['short'] = float(p.extract().replace(",",'.'))
                    count = count+1
                elif count == 3:
                    # d_sub["Total"] = p.extract()
                    asset['total'] = float(p.extract().replace(",",'.'))
                    count = 0
                    repartition[actif.lower()] = dict(asset)
                    asset.clear()
                    # d_sub = {}
        data['repartition_actifs']=dict(repartition)

    def get_perf(self,response,data):
        # Valorisation du fond
        table = response.xpath('//*[@id="overviewQuickstatsDiv"]/table')
        vl_dict = {'devise':'','vl':0}
        raw_vl = table.css('tr')[1].css('td')[2].css('td::text')[0].extract().replace(",",'.')
        vl_dict['devise'] = raw_vl.split('\xa0')[0]
        vl_dict['vl'] = float(raw_vl.split('\xa0')[1])
        data['vl'] = dict(vl_dict)

        # sub_d = {}
        perf = PerfItem()
        table = response.xpath('//*[@id="overviewTrailingReturnsDiv"]/table')

        perf["date"] = table.css('tr')[0].css('td')[1].css('td::text')[0].extract()
        # sub_d["Date"] = table.css('tr')[0].css('td')[1].css('td::text')[0].extract()
        keys = ['ytd','threeY','fiveY','tenY']
        for i in range(1,5):
            # period = table.css('tr')[i].css('td')[0].css('td::text')[0].extract()
            period=keys[i-1]
            value = table.css('tr')[i].css('td')[1].css('td::text')[0].extract().replace(",",'.')
            # sub_d[period] = value
            perf[period]=value

        data['performance'] = dict(perf)

    def get_exposition(self,response,data):
        regions = GeoItem()
        secteurs = SecteurItem()

        try:
            table = response.xpath('//*[@id="overviewPortfolioTopRegionsDiv"]/table')
            for i in range(1,6):
                country = table.css('tr')[i].css('td')[0].css('td::text')[0].extract()
                weight = table.css('tr')[i].css('td')[1].css('td::text')[0].extract()
                regions[format_string(country)]= float(weight.replace(",",'.'))
                # sub_d[str(i)]= country + " - " + weight.replace(",",'.') + " %"

            # Données manquantes dans 'autres'
            sum = 0
            for country in regions.keys():
                sum = sum + float(regions[country])
            regions['autres'] = 100 - sum

            # d['Top Regions'] = sub_d
            data['repartition_geo']= dict(regions)
            regions.clear()
            # sub_d = {}
        except IndexError:
            # d['Top Regions'] = {"1":'-',"2":'-',"3":'-',"4":'-',"5":'-'}
            data['repartition_geo'] = dict(regions)
            regions.clear()

        try:
            table = response.xpath('//*[@id="overviewPortfolioTopSectorsDiv"]/table')
            for i in range(1,6):
                sector = table.css('tr')[i].css('td')[1].css('td::text')[0].extract()
                weight = table.css('tr')[i].css('td')[2].css('td::text')[0].extract()
                secteurs[format_string(sector)]= float(weight.replace(",",'.'))
                # sub_d[str(i)]= sector + " - " + weight.replace(",",'.') + " %"

            # Données manquantes dans 'autres'
            sum = 0
            for sector in secteurs.keys():
                sum = sum + float(secteurs[sector])
            secteurs['autres'] = 100 - sum

            # d['Top Secteurs'] = sub_d
            data['repartition_secteurs'] = dict(secteurs)
            secteurs.clear()
        except IndexError:
            # d['Top Secteurs'] = {"1":'-',"2":'-',"3":'-',"4":'-',"5":'-'}
            data['repartition_secteurs'] = dict(secteurs)
            secteurs.clear()


    def parse(self, response):
        fund = FundItem()
        data = DataItem()
        data['timestamp'] = int(time.time())
        self.get_info(response,fund)
        self.get_allocation(response,data)
        self.get_perf(response,data)
        self.get_exposition(response,data)
        fund['data']=dict(data)
        return fund
