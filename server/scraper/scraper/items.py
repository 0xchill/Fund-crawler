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

class PerfItem(Item):
    date=Field()
    ytd=Field()
    threeY=Field()
    fiveY=Field()
    tenY=Field()

class TopItem(Item):
    first=Field()
    second=Field()
    third=Field()
    fourth=Field()
    fifth=Field()

class FundItem(Item):
    nom=Field()
    vl=Field()
    msCode=Field()
    isin=Field()
    actions=Field()
    obligations=Field()
    liquidites=Field()
    autres=Field()
    performance=Field()
    regions=Field()
    secteurs=Field()
