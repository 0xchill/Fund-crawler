# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# https://doc.org/en/latest/topics/items.html

from scrapy import Field,Item

class AssetItem(Item):
    type=Field()
    long=Field()
    short=Field()
    total=Field()

class TypeItem(Item):
    actions=Field()
    obligations=Field()
    liquidites=Field()
    autres=Field()

class PerfItem(Item):
    date=Field()
    ytd=Field()
    threeY=Field()
    fiveY=Field()
    tenY=Field()

class SecteurItem(Item):
    industriels=Field()
    technologie=Field()
    services_financiers=Field()
    consommation_cyclique=Field()
    consommation_defensive=Field()
    sante=Field()
    materiaux_de_base=Field()
    services_de_communication=Field()
    immobilier=Field()
    services_publics=Field()
    energie=Field()
    autres=Field()

class GeoItem(Item):
    eurozone=Field()
    royaume_uni=Field()
    europe_sauf_euro=Field()
    europe_emergente=Field()
    etats_unis=Field()
    amerique_latine=Field()
    japon=Field()
    asie_emergente=Field()
    asie_pays_developpes=Field()
    canada=Field()
    afrique=Field()
    autres=Field()

class DataItem(Item):
    timestamp=Field()
    vl=Field()
    performance=Field()
    repartition_actifs=Field()
    repartition_geo=Field()
    repartition_secteurs=Field()

class FundItem(Item):
    nom=Field()
    msCode=Field()
    isin=Field()
    data=Field()
