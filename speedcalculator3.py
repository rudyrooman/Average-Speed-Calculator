# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import json
import urllib2
import datetime
import xlsxwriter
import numpy
import math
import pandas as pd


def kmtijd_str_in_sec(kmtijdinmmss):
    """convert timestring in m'ss" to seconds"""
    seconds = int(kmtijdinmmss[-3:-1])
    minutes = int(kmtijdinmmss[:-4])
    return minutes * 60 + seconds


def sec_in_dec(average):
    """ convert time in seconds to minutes decimal notation"""
    return round(float(average) / 60, 2)


def dec_in_mmss_str(average):
    """ convert time in minutes in decimal notation to a string in m'ss"
    """
    return '%s\'%s\"' % (str(math.trunc(average)), str(int((average - math.trunc(average)) * 60)))


def datumhelga(text):
    """ convert date string to dateformat"""
    jaar = int(text[:4])
    maand = int(text[5:7])
    dag = int(text[8:10])
    return datetime.date(jaar, maand, dag)


def refdatum(interval):
    """ interval in year
    # refdatum =datumvandaag - interval
    # yyyy-mm-dd format """
    previousyear = int(datetime.date.today().strftime("%Y")) - interval
    thismonth = int(datetime.date.today().strftime("%m"))
    thisday = int(datetime.date.today().strftime("%d"))
    return datetime.date(previousyear, thismonth, thisday)


def bereken(lijst):
    """ bereken gemiddelde en standarddeviate en verwijder dan alle outliers ,resultaat in min decimale notatie"""
    initial_length = len(lijst)
    lijst.sort()
    # berekenen van gemiddelde snelheid in sec
    average = round(numpy.mean(lijst))
    # berekenen van standard dev in sec
    standarddev = round(numpy.std(lijst))
    # verwijderen van waarden > average + 2 standdard dev
    while lijst[len(lijst) - 1] > average + 2 * standarddev:
        lijst.pop()
    # verwijderen van waarden < average - 2 standdard dev
    lijst.reverse()
    while lijst[len(lijst) - 1] < average - 2 * standarddev:
        lijst.pop()
    print 'We houden rekening met %s/%s km-tijden' % (len(lijst), initial_length)
    # berekenen van gemiddelde snelheid in sec na opkuisen lijst
    average = round(numpy.mean(lijst))
    # berekenen van standard dev in sec na opkuisen lijst
    standarddev = round(numpy.std(lijst))
    return sec_in_dec(average), sec_in_dec(standarddev)


""" inlezen vanuit excel maken want txt werkt alleen in Linux"""
#  inlezen
lopers_namen = {}
lopers_kmtijd = {}
lopers_stdev = {}

# lees gegevens in
df = pd.ExcelFile("lijsthelga.xlsx", encoding='utf8')
# neem het 1de blad als object
sheet1 = df.parse(0)
for teller in range(len(sheet1)):
    row1 = sheet1.iloc[teller].real
    (ID, name) = row1[0], row1[1]
    print ID, name
    lopers_namen[ID] = name

# onderstaande commentaar wegdoen om 1 record te testen
# lopers_namen = {'5336':'Wens Michel'}

for ID in lopers_namen:
    # initialize lijst met km tijden in sec
    lijst = []
    url = 'https://helga-o.com/webres/ws-runner.php?runner=' + str(ID)
    gegevens = urllib2.urlopen(url).read()
    obj = json.loads(gegevens)
    # update exacte naam vanuit Helga
    lopers_namen[ID] = obj.get(u'Name')
    for eventnr in obj[u'Events']:
        event = obj[u'Events'].get(eventnr)
        eventdatum = datumhelga(event.get(u'Date'))
        if eventdatum > refdatum(1):
            eventstatus = event[u'Results'][0].get(u'Status')
            if eventstatus == 'OK':
                eventkmtijd = event[u'Results'][0].get(u'minKm')
                if eventkmtijd != None:
                    lijst.append(kmtijd_str_in_sec(eventkmtijd))
    if len(lijst) >= 2:
        print '%s:' % (lopers_namen[ID]),
        lopers_kmtijd[ID], lopers_stdev[ID] = bereken(lijst)
        print 'De gemiddelde snelheid = %s met een standaardafwijking van %s.' % (
        dec_in_mmss_str(lopers_kmtijd[ID]), dec_in_mmss_str(lopers_stdev[ID]))
        print

# nu alles nog bewaren in excel
# Create a workbook and add a worksheet to save results
workbook = xlsxwriter.Workbook('Output.xlsx')
worksheet = workbook.add_worksheet()
# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0
# Write headers
worksheet.write(row, col, 'Helga-nummer')
worksheet.write(row, col + 1, 'Gem+std')
worksheet.write(row, col + 2, 'Gemiddelde snelheid')
worksheet.write(row, col + 3, 'Standaard dev.')
worksheet.write(row, col + 4, 'Naam')

row = 1
# Iterate over the data and write it out row by row.
for ID in lopers_kmtijd:
    worksheet.write(row, col, int(ID))
    worksheet.write(row, col + 1, lopers_kmtijd[ID] + lopers_stdev[ID])
    worksheet.write(row, col + 2, lopers_kmtijd[ID])
    worksheet.write(row, col + 3, lopers_stdev[ID])
    worksheet.write_string(row, col + 4, lopers_namen[ID])
    row += 1

workbook.close()
