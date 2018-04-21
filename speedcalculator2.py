
from bs4 import BeautifulSoup
import webbrowser
import urllib
import urllib2
import xlsxwriter
import numpy
import re

# lijst inlezen
loperslijst = {}
with open("lijstsnelheid.txt",'rU') as f:
    for line in f:
       (runner, name) = line.split('\t')
       loperslijst[runner] = str(name.rstrip('\n'))
       
# onderstaande commentaar wegdoen om 1 record te testen
#loperslijst = {'44538':'Jespers Leen'}

# resultaatlijst initialiseren
resultaatslijst= {}
standarddevlijst ={}

# re profiles
reaantal= r'\(\d*\)'
rekmtijd= r'\d*\'\d*\"'


# opzoeken loper en berekenen gemiddelde snelheid
for loper in loperslijst:
    print loperslijst[loper]
    url = 'http://helga-o.com/webres/showrunner.php?runner='+loper
    # inlezen website helga
    soup = BeautifulSoup(urllib2.urlopen(url).read(), 'html.parser')
    alletext= soup.get_text()
    aantalmethaakjes=re.findall(reaantal,alletext)[0]
    aantalwed= int(aantalmethaakjes[1:-1])
    kmtijden= re.findall(rekmtijd,alletext)
    kmlijst=[]
    for kmtijd in kmtijden:
       kmlijst.append(kmtijd[2:-1])
    print 'aantal wed',aantalwed
    ftr = [60,1]
    lijst= []
    for it in kmlijst:
         snelheidinsec= sum([a*b for a,b in zip(ftr, map(int,it.split("""\'""")))])
         lijst.append(snelheidinsec)       
    if len(lijst) >=2:
       lijst.sort()
       # berekenen van gemiddelde snelheid in sec
       average= round(numpy.mean(lijst))
       # berekenen van standard dev in sec
       standarddev= round(numpy.std(lijst))
       # verwijderen van waarden > average + 2 standdard dev 
       while lijst[len(lijst)-1]> average+ 2* standarddev :
           lijst.pop()
       # verwijderen van waarden < average - 2 standdard dev 
       lijst.reverse()
       while lijst[len(lijst)-1]< average- 2* standarddev :
           lijst.pop()
       print 'we houden rekening met ', len( lijst),' km-tijden'
       # berekenen van gemiddelde snelheid in sec na opkuisen lijst
       average= round(numpy.mean(lijst))
       #print average
       # berekenen van standard dev in sec na opkuisen lijst
       standarddev= round(numpy.std(lijst))
       #print standarddev
       print ' De gemiddelde snelheid van '+loperslijst[loper] + '= '+str(int(average/60))+ """'""" +str(int(average-60*int(average/60))) +'''"'''
       resultaatslijst[loper]= round(float(average)/60,2)
       print 
       standarddevlijst[loper]= round(float(standarddev)/60,2)


# Create a workbook and add a worksheet to save results
workbook = xlsxwriter.Workbook('Output.xlsx')
worksheet = workbook.add_worksheet()
# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0
# Write headers
worksheet.write(row, col, 'Helga-nummer')
worksheet.write(row, col+1, 'Gemiddelde snelheid')
worksheet.write(row, col+2, 'Standaard dev.')
worksheet.write(row, col+3, 'Naam')

row=1
# Iterate over the data and write it out row by row.
for item  in resultaatslijst:
    worksheet.write(row, col, int(item))
    worksheet.write(row, col + 1, resultaatslijst[item])
    worksheet.write(row, col + 2, standarddevlijst[item])
    worksheet.write_string(row, col + 3, loperslijst[item])
    row += 1

workbook.close()





