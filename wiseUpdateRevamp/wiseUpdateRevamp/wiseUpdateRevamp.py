import webbrowser
from os.path import expanduser
import os, time, decimal
from datetime import date
import datetime
import urllib2, urllib
import re, subprocess
import codecs, csv, io
from llist import dllist, dllistnode
import winsound
from operator import itemgetter
from bs4 import BeautifulSoup, SoupStrainer

class File:
    def __init__(self, data, sku, category, rowIndex, header):
        self.data = data
        self.sku = sku
        self.category = category
        self.rowIndex = rowIndex
        self.header = header

rootDir = expanduser('~')
downloads = rootDir + '/Downloads/'
newImages = downloads + 'newProductImages/'
wiseUpdateReport = open(downloads + 'wiseUpdateReport.txt', 'w')

#create newImages folder in downloads if one doesnt exist
#else clear folder
if not os.path.exists(newImages):
    os.makedirs(newImages)
else:
    for the_file in os.listdir(newImages):
        file_path = os.path.join(newImages, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception:
            print ("File deletion error")

wiseData = File([], [], [], [], [])
medcoProducts = File([], [], [], [], [])
medcoDescriptions = File([], [], [], [], [])
eagleData = File([], [], [], [], [])
compiledData = File([], [], [], {}, [])

#download medcoFile and save latest version file name
def findLatestCopy(fileName):
    highestCopy = 0
    copyNumber = 0
    medcoFile = 0
    tmpMedcoFile = 0

    for root, dirs, files in os.walk(downloads):
        for file in files:
            if fileName in file:
                tmpMedcoFile = file
                if '(' in tmpMedcoFile:
                    buffer = tmpMedcoFile.split('(')
                    copyNumber = buffer[1].split(')')
                    copyNumber = copyNumber[0]
                    if int(copyNumber) > int(highestCopy):
                        highestCopy = copyNumber
                        medcoFile = tmpMedcoFile
    if tmpMedcoFile == 0:
        return -1, 0
    if highestCopy == 0:
        return 0, tmpMedcoFile
    else:
        return highestCopy, medcoFile
def downloadMedcoFile(fileName):
    print('download medcoFile:')
    webbrowser.open("http://liberty.medcocorp.com/marketing/denlorstoolsd35877/medcoproductsd35877.csv")
    highestCopy = 0
    copyThreshold, medcoFile = findLatestCopy(fileName)

    if copyThreshold != -1:
        
        while highestCopy <= copyThreshold:
            highestCopy, medcoFile =  findLatestCopy(fileName)
            print 'waiting for download'
            if '.crdownload' in medcoFile:
                print 'beginning download'
                while '.crdownload' in medcoFile:
                    #medcoFile = medcoFile.replace('.crdownload', '')                  
                    unused, medcoFile = findLatestCopy(fileName)
                    time.sleep(4)
                    print'.',              
            time.sleep(10)           
        return medcoFile
    else:
        while medcoFile == 0:
            highestCopy, medcoFile =  findLatestCopy(fileName)
            time.sleep(10)
            print 'waiting for download'
        return medcoFile
def findLatestWiseFile(fileName):
   
    latestVersion = '0'
    latestDateTime = datetime.datetime(2000, 1, 1, 0, 0, 0)

    currentYear = date.today().year
    for root, dirs, files in os.walk(downloads):
        for file in files:
            if fileName in file:
                nameBuffer = file.replace('.', '-')
                nameBuffer = nameBuffer.split('-')
                nameBuffer[1] = int(nameBuffer[1])
                nameBuffer[2] = int(nameBuffer[2])
                nameBuffer[3] = int(nameBuffer[3])
                
                bufferDateTime = datetime.datetime(nameBuffer[1], nameBuffer[2], nameBuffer[3], 0, 0, 0)
                #print 'buff: ', bufferDateTime

                if bufferDateTime > latestDateTime:
                    latestDateTime = bufferDateTime
                    latestVersion = file
                    #print 'late: ', latestDateTime, '\n'
    return latestVersion
def saveMedcoFeaturesAndBenefits():
    print '\nsave medco features and benefits'
    data = urllib2.urlopen('http://liberty.medcocorp.com/marketing/denlorstoolsd35877/medcoproductsfeaturesbenefitshtml.txt')
    for line in data:
        medcoDescriptions.data.append(re.sub(r'<.*?>', '', line).rstrip().split('|'))
    print 'done'
def saveCsvData(path, fileObj):
    with codecs.open(path, 'r') as f:
        fileObj.data = list(csv.reader(f))
def fileManager():

    medcoFileName = downloadMedcoFile('medcoproducts')
    #medCopy, medcoFileName = findLatestCopy('medcoproducts')#replace with downloadMedcoFile for final version
    print '\n\nmedco file: ', medcoFileName
    wiseFileName =  findLatestWiseFile('products-')
    print 'wise file: ', wiseFileName
    copyNumber, eagleFileName = findLatestCopy('I6000')
    print 'eagle file: ', eagleFileName
    #takes a path and File object
    saveMedcoFeaturesAndBenefits()
    saveCsvData(downloads+wiseFileName, wiseData)
    saveCsvData(downloads+eagleFileName, eagleData)
    saveCsvData(downloads+medcoFileName, medcoProducts)
def testPrintCsvRead():
     wiseFileName =  findLatestWiseFile('products-')
     saveCsvData(downloads+wiseFileName, wiseData)
def printToCsv():
    

    def removeUnusedData():
        cleanRow = []
        cleanedData = []
        for rNum, row in enumerate(wiseData.data):
            for cNum, col in enumerate(row):
                if(cNum in wiseHeadersUsed):
                    #del wiseData.data[rNum][cNum]
                    cleanRow.append(wiseData.data[rNum][cNum])
            cleanedData.append(cleanRow)
            cleanRow = []
        return cleanedData

    global cleanedData
    global cleanRow
    print '\nwriting data'
    count = 0
    fileCount = 0
    #writer = csv.writer(open(downloads + 'updatedWiseFile-1.csv', 'wb'))
    wiseQtrLength = int(len(wiseData.data) / 4) + 1 

    cleanedData = removeUnusedData()

    for row in cleanedData:
        #make new file for file splitting
        if(count % wiseQtrLength == 0):
                fileCount+=1
                writer = csv.writer(open(downloads + 'updatedWiseFile-'+str(fileCount)+'.csv', 'wb'))
                if(count != 0):
                    writer.writerow(cleanedData[0])
        count+=1
        try:
            writer.writerow(row).decode('utf8')
        except:
                ''     
    print 'done'
def makeList():
    #wise
    wiseData.header = []
    for i in wiseData.data[0]:
         wiseData.header.append(i)

    #medco
    medcoProducts.header = []
    for i in medcoProducts.data[0]:
        medcoProducts.header.append(i)

    medcoSku = medcoProducts.header.index('SKU')
    medcoProducts.sku = []
    for i in medcoProducts.data:
        medcoProducts.sku.append(i[medcoSku])
    #medco descriptions
    medcoDescriptions.sku = []
    for i in medcoDescriptions.data:
        medcoDescriptions.sku.append(i[0])
    #eagle
    eagleData.header = []
    for i in eagleData.data[0]:
        eagleData.header.append(i)
    #compiledData
    compiledData.header = []
    for i in compiledData.data[0]:
        compiledData.header.append(i)
        
def makeCompiledData():
    print '\ncombine update files'
    compiledData.data.append([])

    compiledData.data[0].extend(('Product Name', 'Product Code/SKU', 'Brand Name', 'Product Description', 'Price', 'Cost Price', 'Product Weight', 'Product Width', 'Product Height', 'Product Depth', 'Allow Purchases?', 'Product Visible?', 'Product Availability', 'Category', 'Product Image File - 1', 'Product Image Description - 1', 'Product Image Is Thumbnail - 1', 'Product URL', 'map'))
    
    makeList()
    #compiledData indices
    compiledData.rowIndex['compProductName'] = compProductName = compiledData.header.index('Product Name')    
    compiledData.rowIndex['compSku'] = compSku = compiledData.header.index('Product Code/SKU')
    compiledData.rowIndex['compBrand'] = compBrand = compiledData.header.index('Brand Name')
    compiledData.rowIndex['compDesc'] = compDesc = compiledData.header.index('Product Description')
    compiledData.rowIndex['compPrice'] = compPrice = compiledData.header.index('Price')
    compiledData.rowIndex['compCost'] = compCost = compiledData.header.index('Cost Price')
    compiledData.rowIndex['compWeight'] = compWeight = compiledData.header.index('Product Weight')
    compiledData.rowIndex['compWidth'] = compWidth = compiledData.header.index('Product Width')
    compiledData.rowIndex['compHeight'] = compHeight = compiledData.header.index('Product Height')
    compiledData.rowIndex['compDepth'] = compDepth = compiledData.header.index('Product Depth')
    compiledData.rowIndex['compAllow'] = compAllow = compiledData.header.index('Allow Purchases?')
    compiledData.rowIndex['compVisible'] = compVisible = compiledData.header.index('Product Visible?')
    compiledData.rowIndex['compAvailable'] = compAvailable = compiledData.header.index('Product Availability')
    compiledData.rowIndex['compCategory'] = compCategory = compiledData.header.index('Category')
    compiledData.rowIndex['compImage'] = compImage = compiledData.header.index('Product Image File - 1')
    compiledData.rowIndex['compImageDes'] = compImageDes = compiledData.header.index('Product Image Description - 1')
    compiledData.rowIndex['compThumb'] = compThumb = compiledData.header.index('Product Image Is Thumbnail - 1')
    compiledData.rowIndex['compUrl'] = compUrl = compiledData.header.index('Product URL')
    compiledData.rowIndex['compMap'] = compMap = compiledData.header.index('map')

    #medco indices
    compiledWidth = len(compiledData.data[0])
    medcoSku = medcoProducts.header.index('SKU')
    medcoProductName = medcoProducts.header.index('ProductName')
    medcoBrand = medcoProducts.header.index('MfgName')
    medcoDesc = medcoProducts.header.index('FullProductName')
    medcoCost = medcoProducts.header.index('Cost')
    medcoWeight = medcoProducts.header.index('Weight')
    medcoHazmat = medcoProducts.header.index('HazmatInd')
    medcoTruck = medcoProducts.header.index('TruckInd')
    medcoPublish = medcoProducts.header.index('REASON_CD')
    medcoUrl = medcoProducts.header.index('ImageUrl')
    medcoAvail = medcoProducts.header.index('Inventory')
    medcoImage = medcoProducts.header.index('ImageUrl')
    medcoMap = medcoProducts.header.index('MAP_PRICE_2')

    #add medco products to compiledData
    for indx, value in enumerate(medcoProducts.data[1:]):
        compiledData.data.append([''] * compiledWidth)

        compiledData.data[-1][compProductName] = value[medcoSku] + ' ' + value[medcoProductName]
        compiledData.data[-1][compSku] = value[medcoSku]
        compiledData.data[-1][compBrand] = value[medcoBrand]
        compiledData.data[-1][compDesc] = 'medco'#value[medcoDesc]

        try:
            float(value[medcoCost])
            compiledData.data[-1][compCost] = value[medcoCost]
        except:
            compiledData.data[-1][compCost] = '0'
            print 'non-numeric price for: ', value[medcoSku]

        compiledData.data[-1][compWeight] = value[medcoWeight]

        if ('NO_HAZMAT' == value[medcoHazmat] and
            'NO_TRUCK' == value[medcoTruck] and
            'PUBLISH' == value[medcoPublish] and
            value[medcoUrl]):

            compiledData.data[-1][compAllow] = 'Y'
            compiledData.data[-1][compVisible] = 'Y'
        else:
            compiledData.data[-1][compAllow] = 'N'
            compiledData.data[-1][compVisible] = 'N'

        compiledData.data[-1][compAvailable] = value[medcoAvail]
        compiledData.data[-1][compCategory] = 'Medco Products;Hand Tools/Open Stock'
        compiledData.data[-1][compImage] = value[medcoSku] + '.JPG'
        compiledData.data[-1][compImageDes] = value[medcoProductName]
        compiledData.data[-1][compThumb] = 'Y'
        compiledData.data[-1][compUrl] = value[medcoImage]
        compiledData.data[-1][compMap] = value[medcoMap]
    
    #eagle indices
    eagleSku = eagleData.header.index('Item Number')
    eagleDes1 = eagleData.header.index('Description 1')
    eagleDes2 = eagleData.header.index('Description 2')
    eagleBrand = eagleData.header.index('Vendor Name')
    eaglePrice = eagleData.header.index('Price')
    eagleWeight = eagleData.header.index('Item Weight')
    eagleWidth = eagleData.header.index('Width')
    eagleHeight = eagleData.header.index('Height')
    eagleLength = eagleData.header.index('Length')
    eagleFreight = eagleData.header.index('Freight Policy')
    eagleAvail = eagleData.header.index('Avail Combined')
    eagleUrl = eagleData.header.index('Image Path')
    eagleMap = eagleData.header.index('Map Price')

    for indx, value in enumerate(eagleData.data[1:]):
        compiledData.data.append([''] * compiledWidth)

        if value[eagleDes1]:
            compiledData.data[-1][compProductName] = value[eagleSku] + ' ' + value[eagleDes1]
        else:
            compiledData.data[-1][compProductName] = value[eagleSku]

        compiledData.data[-1][compSku] = value[eagleSku]
        #print indx
        compiledData.data[-1][compBrand] = value[eagleBrand]

        if value[eagleDes1] and value[eagleDes2]:
            compiledData.data[-1][compDesc] = value[eagleDes1] + ' ' + value[eagleDes2]
        elif value[eagleDes1]:
            compiledData.data[-1][compDesc] = value[eagleDes1]
        elif value[eagleDes2]:
            compiledData.data[-1][compDesc] = value[eagleDes2]

        try:
            float(value[eaglePrice])
            compiledData.data[-1][compCost] = value[eaglePrice]
        except:
            compiledData.data[-1][compCost] = '0'
            print 'non-numeric price for: ', value[eagleSku]

        compiledData.data[-1][compWeight] = value[eagleWeight]
        compiledData.data[-1][compWidth] = value[eagleWidth]
        compiledData.data[-1][compHeight] = value[eagleHeight]
        compiledData.data[-1][compDepth] = value[eagleLength]

        if value[eagleFreight] == 'REG POLICY':
            compiledData.data[-1][compAllow] = 'Y'
            compiledData.data[-1][compVisible] = 'Y'
        else:
            compiledData.data[-1][compAllow] = 'N'
            compiledData.data[-1][compVisible] = 'N'

        compiledData.data[-1][compAvailable] = value[eagleAvail]
        compiledData.data[-1][compCategory] = 'Eagle Products;Hand Tools/Open Stock'
        compiledData.data[-1][compImage] = value[eagleSku] + '.JPG'
        compiledData.data[-1][compImageDes] = value[eagleDes1]
        compiledData.data[-1][compThumb] = 'Y'
        compiledData.data[-1][compUrl] = value[eagleUrl]
        compiledData.data[-1][compMap] = value[eagleMap]

    print 'done\n'
def is_number(str):
    try:
        float(str)
        return True
    except:
        return False
def updatePrice():
    print 'update price'
    colIndex = compiledData.rowIndex

    for indx, value in enumerate(compiledData.data[1:]):
        if value[colIndex['compCost']] and is_number(value[colIndex['compCost']]):
            costPrice = float(value[colIndex['compCost']])
            if costPrice <= 450:
                price = str(round(costPrice * 1.26, 2))
            elif costPrice > 450 and costPrice <= 750:
                price = str(round(costPrice * 1.2, 2))
            elif costPrice > 750 and costPrice <= 999:
                price = str(round(costPrice * 1.15, 2))
            elif costPrice > 999:
                price = str(round(costPrice * 1.1, 2))

            value[colIndex['compPrice']] = price
        else:
            print 'error in file, check: ', value[colIndex['compSku']]
            value[colIndex['compPrice']] = 'null'
    print 'done\n'
def playSong():
    
    def chord(root_frequency):
        winsound.Beep(int(root_frequency), 60)
        winsound.Beep(int(root_frequency*1.25), 60)
        winsound.Beep(int(root_frequency*1.5), 60)
        winsound.Beep(int(root_frequency*2), 100)

    chord(261.626)
    time.sleep(0.35)
    chord(261.626)
    time.sleep(0.05)
    chord(233.082)
    time.sleep(0.35)
    chord(233.082)
    time.sleep(0.05)
    chord(207.652)
    time.sleep(0.35)
    chord(207.652)
    time.sleep(0.05)
    chord(195.998)
    time.sleep(0.35)
    chord(195.998)
    time.sleep(0.05)
def addItemToWise(comp):
    wiseData.data.append([''] * len(wiseData.data[0]))
    
    wiseData.data[-1][wiseProductName] = comp[compiledData.rowIndex['compProductName']]
    wiseData.data[-1][wiseSku] = comp[compiledData.rowIndex['compSku']]
    wiseData.data[-1][wiseBrand] = comp[compiledData.rowIndex['compBrand']]
    wiseData.data[-1][wiseDesc] = comp[compiledData.rowIndex['compDesc']]
    wiseData.data[-1][wiseCost] = comp[compiledData.rowIndex['compCost']]

    compMap = comp[compiledData.rowIndex['compMap']]

    try:
        if float(compMap) != 0:
            wiseData.data[-1][wisePrice] = compMap
        else:
            wiseData.data[-1][wisePrice] = comp[compiledData.rowIndex['compPrice']]
    except:
        wiseData.data[-1][wisePrice] = comp[compiledData.rowIndex['compPrice']]
        print '\nerror'
        print 'non-numeric value for map price: ', comp[compiledData.rowIndex['compSku']]
        print

    wiseData.data[-1][wiseWeight] = comp[compiledData.rowIndex['compWeight']]
    wiseData.data[-1][wiseWidth] = comp[compiledData.rowIndex['compWidth']]
    wiseData.data[-1][wiseHeight] = comp[compiledData.rowIndex['compHeight']]
    wiseData.data[-1][wiseDepth] = comp[compiledData.rowIndex['compDepth']]
    wiseData.data[-1][wiseAllow] = comp[compiledData.rowIndex['compAllow']]
    wiseData.data[-1][wiseVisible] = comp[compiledData.rowIndex['compVisible']]

    try:
        if float(comp[compiledData.rowIndex['compAvailable']]) == 0:
             wiseData.data[-1][wiseAllow] = 'N'
        else:
             wiseData.data[-1][wiseAllow] = 'Y'
    except:
        wiseData.data[-1][wiseAllow] = 'N'

    wiseData.data[-1][wiseAvailable] = comp[compiledData.rowIndex['compAvailable']]
    wiseData.data[-1][wiseCategory] = comp[compiledData.rowIndex['compCategory']]
    wiseData.data[-1][wiseImage] = comp[compiledData.rowIndex['compImage']]
    wiseData.data[-1][wiseImageDes] = comp[compiledData.rowIndex['compImageDes']]
    wiseData.data[-1][wiseThumb] = 'Y'

    tmp = comp[compiledData.rowIndex['compProductName']].replace(' ', '-')
    tmp = re.sub(r'[^a-zA-Z0-9-]', '', tmp)
    wiseData.data[-1][wiseUrl] = '/' + tmp + '/'

    #if description is preset to medco, search medcoDescription for product description
    if comp[compiledData.rowIndex['compDesc']] == 'medco':
        if comp[compiledData.rowIndex['compSku']] in medcoDescriptions.sku:
            medcoDescIndex = medcoDescriptions.sku.index(comp[compiledData.rowIndex['compSku']])
            wiseData.data[-1][wiseDesc] = medcoDescriptions.data[medcoDescIndex][1]
        
    #download and save product image to download folder
    path = newImages + comp[compiledData.rowIndex['compSku']] + '.JPG'
    url = comp[compiledData.rowIndex['compUrl']]

    try:
        urllib.urlretrieve(url, path) 
    except urllib2.URLError, e:
        #print('url error 1')
        #time.sleep(1)
        ''
    except:
        #print('url error 2')
        ''
    #if product has no description or no image it is hidded
    if not wiseData.data[-1][wiseDesc] or wiseData.data[-1][wiseDesc] == 'medco':
        wiseData.data[-1][wiseVisible] = 'N'
        wiseData.data[-1][wiseAllow] = 'Y'
    try:
        if os.path.getsize(path) < 1000:
            os.remove(path)
            wiseData.data[-1][wiseVisible] = 'N'
    except:
        ''
        wiseData.data[-1][wiseVisible] = 'N'
        wiseData.data[-1][wiseAllow] = 'Y'
def getWiseUrl(webPage, sku):
    
    try:
        buff = urllib2.urlopen(webPage).read()

        soup = BeautifulSoup(buff)
        imgTags = soup.findAll('img')

        #img = [tag['src'] for tag in imgTags]
        img = []
        for tag in imgTags:
            img.append(tag['src'])
            if sku in img[-1]:
                return img[-1]
    except:
        return 'false'
def checkItemForUpdatedDescriptionAndImage(comp, wiseIndex):

    path = newImages + comp[compiledData.rowIndex['compSku']] + '.JPG'
    wiseData.data[wiseIndex][wiseImage] = comp[compiledData.rowIndex['compSku']] + '.JPG'
    wiseData.data[wiseIndex][wiseImageDes] = comp[compiledData.rowIndex['compImageDes']]
    wiseData.data[wiseIndex][wiseThumb] = 'Y'

    global oldItemsWithNewDescriptionsOrImages

    #update description if necessary
    if not wiseData.data[wiseIndex][wiseDesc] and comp[compiledData.rowIndex['compDesc']]:
        wiseData.data[wiseIndex][wiseDesc] = comp[compiledData.rowIndex['compDesc']]

    elif  wiseData.data[wiseIndex][wiseDesc] == 'medco':
        if comp[compiledData.rowIndex['compSku']] in medcoDescriptions.sku:
            medcoDescIndex = medcoDescriptions.sku.index(comp[compiledData.rowIndex['compSku']])
            wiseData.data[wiseIndex][wiseDesc] = medcoDescriptions.data[medcoDescIndex][1]

    #if wiseData.data has valid url keep it else get new url from update file
    if wiseData.data[wiseIndex][wiseDesc] and wiseData.data[wiseIndex][wiseDesc] != 'medco':
        if wiseData.data[wiseIndex][wiseImageUrl]:
            
            wiseData.data[wiseIndex][wiseVisible] = 'Y'
            oldItemsWithNewDescriptionsOrImages += 1
        else:
            wiseData.data[wiseIndex][wiseVisible] = 'N'
            url = comp[compiledData.rowIndex['compUrl']]
            try:
                urllib.urlretrieve(url, path)
            except:
                ''
            try:
                if os.path.getsize(path) < 1000:
                    os.remove(path)
                else:                   
                    wiseData.data[wiseIndex][wiseVisible] = 'Y'
                    oldItemsWithNewDescriptionsOrImages += 1
            except:
                ''
def updateWisePriceAndAvailability():
    #linked list holds sku, index in wiseData.data
    print '\nmake linked list'
    wiseSku = wiseData.header.index('Product Code/SKU')
    wiseList = dllist()

    for indx, value in enumerate(wiseData.data[1:]):
        wiseList.appendright([value[wiseSku], indx + 1])

    print 'done\n'
    print 'update wise price and availability and add new items'

    #checks if sku is in wise linked list. if so the item is updated in wiseData.data and is removed from linked list
    #                                      else it is added to wiseData.data
    compSku = compiledData.rowIndex['compSku']
    compCostPrice = compiledData.rowIndex['compCost']
    compPrice = compiledData.rowIndex['compPrice']
    compAvailable = compiledData.rowIndex['compAvailable']
    compMap = compiledData.rowIndex['compMap']
    compVisible = compiledData.rowIndex['compVisible']
    compDesc = compiledData.rowIndex['compDesc']

    newItems = 0
    itemPriceUpdate = 0
    manuallyAdded_hiddenNonListed = 0
    totalItemCount = 0

    for compIndx, compValue in enumerate(compiledData.data[1:]):
        inList = False
        totalItemCount += 1

        if totalItemCount % 10000 == 0:
                print totalItemCount, ' items updated'
        for wiseIndx, wiseValue in enumerate(wiseList):
   
            if wiseValue[0] == compValue[compSku]:
                #print 'true'
                inList = True
                if 'Manually Added' not in wiseData.data[wiseValue[1]][wiseCategory] and 'Hidden non listed' not in wiseData.data[wiseValue[1]][wiseCategory]:
                    wiseData.data[wiseValue[1]][wiseCost] = compValue[compCostPrice]
                    wiseData.data[wiseValue[1]][wisePrice] = compValue[compPrice]
                    wiseData.data[wiseValue[1]][wiseAvailable] = compValue[compAvailable]

                    if (wiseData.data[wiseValue[1]][wiseVisible] == 'N' or not wiseData.data[wiseValue[1]][wiseImage]) and compValue[compVisible] == 'Y' and checkForUpdates == 'y':
                        #updates product visibility
                        checkItemForUpdatedDescriptionAndImage(compValue, wiseValue[1])

                    if float(compValue[compMap]) != 0:
                        wiseData.data[wiseValue[1]][wisePrice] = compValue[compMap]
                    
                    if float(compValue[compAvailable]) == 0:
                        wiseData.data[wiseValue[1]][wiseAllow] = 'N'
                    else:
                        wiseData.data[wiseValue[1]][wiseAllow] = 'Y'

                    itemPriceUpdate += 1
                else:
                    manuallyAdded_hiddenNonListed += 1
                #print wiseList.nodeat(wiseIndx)
                wiseList.remove(wiseList.nodeat(wiseIndx))
                break
            
        if inList == False:
            #print 'false'
            addItemToWise(compValue)
            newItems += 1
            #print 'new Item: ', compValue[compSku]
    countDiscontinued = 0

    print 'done'

    #hide items left in linked list
    print '\nhide discontinued'
    for i in wiseList:
        if 'Manually Added' not in wiseData.data[i[1]][wiseCategory] and 'Hidden non listed' not in wiseData.data[i[1]][wiseCategory]:
            print i[0], ' ', i[1]
            wiseData.data[i[1]][wiseAllow] = 'N'
            wiseData.data[i[1]][wiseVisible] = 'N' 
            countDiscontinued += 1
        else:
            manuallyAdded_hiddenNonListed += 1
    print 'done\n'

    print 'number items discontinued: ', countDiscontinued
    print 'new items added: ', newItems
    print 'manually added or hidden non listed: ', manuallyAdded_hiddenNonListed
    print 'old items with updated descriptions/images: ', oldItemsWithNewDescriptionsOrImages
    print 'number of prices updated: ', itemPriceUpdate

    wiseUpdateReport.writelines(['number items discontinued: ', str(countDiscontinued), '\n'])
    wiseUpdateReport.writelines(['new items added: ', str(newItems), '\n'])
    wiseUpdateReport.writelines(['manually added or hidden non listed: ', str(manuallyAdded_hiddenNonListed), '\n'])
    wiseUpdateReport.writelines(['old items with updated descriptions/images: ', str(oldItemsWithNewDescriptionsOrImages), '\n'])
    wiseUpdateReport.writelines(['number of prices updated: ', str(itemPriceUpdate), '\n'])
def merge(left, right, sortIndex):
    result = []

    while left or right:
        if left and right:
            if float(left[0][sortIndex]) <= float(right[0][sortIndex]):
                result.append(left[0])
                del left[0]
            else:
                result.append(right[0])
                del right[0]
        elif left:
            result.append(left[0])
            del left[0]
        elif right:
            result.append(right[0])
            del right[0]
    return result
def merge_sort(data, sortIndex):
    if len(data) <= 1:
        return data

    left = []
    right = []

    middle = len(data)/2

    for i in range(middle):
        left.append(data[i])
    for i in range(middle, len(data)):
        right.append(data[i])

    left = merge_sort(left, sortIndex)
    right = merge_sort(right, sortIndex)

    return merge(left, right, sortIndex)
def sortManager():
    print 'sorting data'

    firstRow = compiledData.data[0]
    result = merge_sort(compiledData.data[1:], 5)

    compiledData.data = [firstRow] + result

    firstRow = wiseData.data[0]
    result = merge_sort(wiseData.data[1:], wiseCost)

    wiseData.data = [firstRow] + result
    print 'done'
def printUpdatedList():
    print '\nwriting data'
    count = 0;
    updatedWiseFile = open(downloads + 'updatedWiseFile.txt', 'w')
    #writer = csv.writer(file, delimiter = ',', lineterminator = '\n')

    for row in wiseData.data:
        for col in wiseHeadersUsed:
            try:
                updatedWiseFile.write(row[col]).decode('utf8')
            except:
                ''
            updatedWiseFile.write('\t')
        updatedWiseFile.write('\n')
    print 'done'


oldItemsWithNewDescriptionsOrImages = 0
checkForUpdates = 'z'

print 'Do you want to check non-visible items for updates? y(slow), n(fast)'

while checkForUpdates != 'n' and checkForUpdates != 'y':
    checkForUpdates = raw_input('Enter: y or n\n')
    if checkForUpdates != 'n' and checkForUpdates != 'y':
        print '\ninvalid input'
print
fileManager()
#testPrintCsvRead()

makeCompiledData()

#adapt headers for unique file headers
#makeList()
for i in wiseData.data[0]:
    #for new headers
    if i == 'Code':

        wiseData.data[0][wiseData.header.index('Name')] = 'Product Name'
        wiseData.data[0][wiseData.header.index('Code')] = 'Product Code/SKU'
        wiseData.data[0][wiseData.header.index('Brand')] = 'Brand Name'
        wiseData.data[0][wiseData.header.index('Description')] = 'Product Description'
        #cost price is same
        wiseData.data[0][wiseData.header.index('Calculated Price')] = 'Price'
        wiseData.data[0][wiseData.header.index('Weight')] = 'Product Weight'
        wiseData.data[0][wiseData.header.index('Width')] = 'Product Width'
        wiseData.data[0][wiseData.header.index('Height')] = 'Product Height'
        wiseData.data[0][wiseData.header.index('Depth')] = 'Product Depth'
        wiseData.data[0][wiseData.header.index('Allow Purchases')] = 'Allow Purchases?'
        wiseData.data[0][wiseData.header.index('Product Visible')] = 'Product Visible?'
        #product availability is same
        wiseData.data[0][wiseData.header.index('Category Details')] = 'Category'
        wiseData.data[0][wiseData.header.index('Images')] = 'Product Image File - 1'
        wiseData.data[0][wiseData.header.index('Page Title')] = 'Product Image Description - 1'
        wiseData.data[0][wiseData.header.index('META Keywords')] = 'Product Image Is Thumbnail - 1'
        #product url is same
        
        
        makeList()

wiseProductName = wiseData.header.index('Product Name')
wiseSku = wiseData.header.index('Product Code/SKU')
wiseBrand = wiseData.header.index('Brand Name')
wiseDesc = wiseData.header.index('Product Description')
wiseCost = wiseData.header.index('Cost Price')
wisePrice = wiseData.header.index('Price')
wiseWeight = wiseData.header.index('Product Weight')
wiseWidth = wiseData.header.index('Product Width')
wiseHeight = wiseData.header.index('Product Height')
wiseDepth = wiseData.header.index('Product Depth')
wiseAllow = wiseData.header.index('Allow Purchases?')
wiseVisible = wiseData.header.index('Product Visible?')
wiseAvailable = wiseData.header.index('Product Availability')
wiseCategory = wiseData.header.index('Category')
wiseImage = wiseData.header.index('Product Image File - 1')
wiseImageDes = wiseData.header.index('Product Image Description - 1')
wiseThumb = wiseData.header.index('Product Image Is Thumbnail - 1')
wiseUrl = wiseData.header.index('Product URL')
wiseImageUrl = wiseData.header.index('Product Image URL - 1')


#for printing to file
wiseHeadersUsed = [wiseProductName, wiseSku, wiseBrand, wiseDesc, wisePrice, wiseCost, wiseWeight, wiseWidth, wiseHeight, wiseDepth, wiseAllow, wiseVisible, wiseAvailable, wiseCategory, wiseImage, wiseImageDes, wiseThumb, wiseUrl]


updatePrice()
sortManager()
updateWisePriceAndAvailability()


printUpdatedList()
printToCsv()
os.startfile(newImages)
#playSong()

print '\nupdate complete'



