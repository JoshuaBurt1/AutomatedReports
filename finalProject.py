#TRANSPORTATION FLEET COST CALCULATOR AND REPORT GENERATOR
import pyinputplus as pyip
import ezsheets, random #module allow your program to log in to Google’s servers and make API requests. 
import urllib.request, re
import os, docx, subprocess, datetime, time

#Input validation: inputYesNo will only allow yes/no input; will output lower case if enter in capitals
print("Energy Cost Report")
fleetSize = pyip.inputInt(prompt = "Enter the your transport fleet size. Enter a number. ", blank = False, lessThan=11)
print(fleetSize)
customerNumber = pyip.inputInt(prompt = "Enter the number of customers per transport. Enter a number. ", blank = False, lessThan=50)
print(customerNumber)
reportChoice = pyip.inputYesNo(prompt = "Does the user want to pause automated report writing? Choices: yes/no. ", blank = False, blockRegexes = ['\d'])
print(reportChoice)
if reportChoice == "yes": #view report; having this open halts docx update
        subprocess.Popen(["C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE","C:\\Users\\Josh\\Desktop\\Python\\CourseCOMP1112\\FinalProject\\DailyTransportExpenses.docx"]) 

while reportChoice != "yes":
    #GOOGLE SHEETS FORMATTING VARIABLES 
    googleSheet = ezsheets.createSpreadsheet(title='Energy Cost Report') #this creates a Google Sheet
    gs = googleSheet.sheets[0]
    letterList = ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"]
    numberList = ["1","2","3","4","5","6","7","8","9","0"]
    transportNameRow = 3 
    custGoogleSheet = 3
    transportCoordRow = 4 
    totalRowFormat = 4
    distanceDifRow = 9
    totalColFormat = 12
    gs.updateRow(1, ["Transport Energy Cost Estimator"])
    gs.updateRow(2, ["Transport Name, Satellite Distribution Center Coordinates, Distance from Home Office [0,0]"])
    gs.updateRow(7, ["Customer, Coordinate of Customer, Distance between Customer and transport"])

    #PART 1: INITIAL DATA GENERATION
    def nameGen(): #Generates a list of names based on user input (fleetsize)
        name = ""
        randomLength = random.randint(4,8)
        for letter in range(randomLength): #single name creation
            randomLetter = random.randint(0,25)
            name+=(letterList[randomLetter]) #creates a list of characters of length 4 to 8
        capital=name[0].upper() #capitalizes first letter
        newName=capital+name[1:len(name)]
        return newName
    
    def nameList(): #create a list of transport and customer names
        names = []
        for nameX in range(fleetSize): #name list creation
            word = nameGen() #obtains a different name for each list element
            names.append(word) #creates a list of names
        return names #calling this function returns a list of names

    #Coordinate pair generation
    def coordGen(): #Generates a list of coordinate pairs (ie. [[01,23],[04,56]]) based on user input (fleetsize)
        xCoord = ""
        randomLength = random.randint(1,2)
        for numberX in range(randomLength):
            randomNumber = random.randint(0,9)
            xCoord+=(numberList[randomNumber]) #creates x-coordinate (1 to 2 digits in length ie. 01)
        coordIntermediary=xCoord+"," #appends comma to x-coordinate (ie. 01,)
        yCoord = ""
        randomLength = random.randint(1,2)
        for numberY in range(randomLength):
            randomNumber = random.randint(0,9)
            yCoord+=(numberList[randomNumber]) #creates y-coordinate (1 to 2 digits in length ie. 23)
        coordinates=coordIntermediary+yCoord #creates a coordinate pair (01,23)
        cordFormatInt=coordinates.rjust(len(coordinates)+1,"[")
        coordFormatFinal=cordFormatInt.ljust(len(cordFormatInt)+1,"]") #creates a coordinate pair [01,23]
        return coordFormatFinal
    
    def coordList(): #creates a list of [x,y] coordinates
        distances = [] 
        for coordXY in range(fleetSize): #coordList creation
            coordinates = coordGen() #obtains a different coordinate pair for each list element
            distances.append(coordinates) #creates a list of coordinate pairs [[01,23],[04,56]]
        return distances #calling this function returns a list of coordinate pairs

    #PART 2: EXTRACT DATA FROM GOOGLE SHEETS TO PERFORM DISTANCE CALCULATION
    def squareRoot(zSquare): #calculates square root 
        if zSquare < 2: #Example C.: zSquare = 4
            return zSquare #if zSquare = 1, square root is 1
        zSquare2 = zSquare #Example C.: zSquare2 = 4
        z = (zSquare2 + 1)/2 #Example C.: z = (4+1)/2 = 5/2 = 2.5
        while abs(zSquare2-z) >0.1: #controls exactness of z (square root value); 0.001 is more exact than 0.1
            zSquare2 = z #Example C.: zSquare2 = 5/2 =2.5
            z = (zSquare2+(zSquare/zSquare2))/2 #First iteration example: z = ((n+1)/2+(n/((n+1)/2)))/2; A. n=16-->5.19; B. n=9-->3.4; C. n=4-->2.05
        #next while loop: 2.5-2.05 > 0.1
        #zSquare2 = 2.05
        #z = (2.05+(4/2.05))/2 = 2.0006
        #next while loop: 2.05-2.0006 !> 0.1 --> loop stops
        return z #calling this function returns a z value (minimum distance between coordinates)
    
    def transportDistanceCalc(): #Calculates transport vehicle minimum distance to Home Office (0,0)
        transpCoord = gs.getRow(transportCoordRow) #1. Extract data from gs (sheet2)
        zDistance = []
        global transportCoords #so variable can be used in customerDistance()
        transportCoords = []
        for tCell, value in enumerate(transpCoord): #cleanse data of non-integer symbols
            if len(transpCoord[tCell]) > 0:
                transpCoord[tCell] = str(transpCoord[tCell]).lstrip('[')
                transpCoord[tCell] = str(transpCoord[tCell]).rstrip(']')
                transpCoord[tCell] = str(transpCoord[tCell]).split(',')
                x = int(transpCoord[tCell][0]) #difference between transport and home office (x)
                y = int(transpCoord[tCell][1]) #difference between transport and home office (y)
                zSquare = (x*x+y*y) #Transport to home office distance^2 (x,y)
                zDistance.append(squareRoot(zSquare)) #creates a list of distance numbers of length 1 to 2 digits
                transportCoords.append(transpCoord[tCell])
        return zDistance #returns a list of transport to home office distances

    def rowCoord(): #changes the row formatting for Google Sheets customerDistance() calculation placement
        global distanceDifRow
        distanceDifRow+=3
        return distanceDifRow
    
    def customerDistance(): #Calculates transport vehicle minimum distance to customers (x,y)
        customerCoord = gs.getRow(distanceDifRow) 
        z2Distance = []
        custCoords = []
        for cCell, value in enumerate(customerCoord): #cleanse data of non-integer symbols
            if len(customerCoord[cCell]) > 0:
                customerCoord[cCell] = str(customerCoord[cCell]).lstrip('[')
                customerCoord[cCell] = str(customerCoord[cCell]).rstrip(']')
                customerCoord[cCell] = str(customerCoord[cCell]).split(',')
                xB = int(customerCoord[cCell][0])
                yB = int(customerCoord[cCell][1])
                xC = abs(int(transportCoords[cCell][0])-xB) #absolute difference between transport and customer (x)
                yC = abs(int(transportCoords[cCell][1])-yB) #absolute difference between transport and customer (y)
                zSquare2 = (xC*xC+yC*yC) #Transport to customer distance^2 (x,y)
                z2Distance.append(squareRoot(zSquare2)) #creates a list of distance numbers of length 1 to 2 digits
                custCoords.append(customerCoord[cCell])
        return z2Distance #returns a list of transport to customer distances

    #PART 3: TRANSPORT VEHICLE AND CUSTOMER WRITTEN TO GOOGLE SHEETS
    def dataInputTransport(): #generate transport location data
        for rowT in range(3,4): #loops through row 1 to 2; must start at row 1
            gs.updateRow(rowT, nameList()) #updates each row with a list of 10 transport names
        for rowT in range(4,5): #loops through row 1 to 2; must start at row 1
            gs.updateRow(rowT, coordList()) #updates each row with a list of 10 transport names
        for rowT in range(5,6): #AFTER EXTRACTION FROM SHEET 1, MUST BE IN THIS LOCATION
            gs.updateRow(rowT, transportDistanceCalc()) 
    dataInputTransport()
    
    def dataInputCustomer(): #generate customer location data based on user input (customerNumber)
        for step3 in range(0,customerNumber*custGoogleSheet,3): #m increase by +3 every iteration (0,3,6..)
            for rowC in range(8+step3,9+step3): 
                gs.updateRow(rowC, nameList()) 
            for rowC in range(9+step3,10+step3): 
                gs.updateRow(rowC, coordList())
            for rowC in range(10+step3,11+step3):
                gs.updateRow(rowC, customerDistance())
            rowCoord() #changes row value so that customerDistance reads the correct coordinates (every 3rd row) 
    dataInputCustomer()

    #PART 4: WEB SCRAPE CURRENT GAS PRICE 
    webUrl = urllib.request.urlopen("http://stockr.net/toronto/gasprice.aspx")
    data=webUrl.read()
    regex = rb'(5\d|[6-9]\d|[12]\d{2}|300)' #search html code for gas range 50 to 300
    pattern  = re.compile(regex)
    gas = re.findall(pattern, data) #searches regex
    gas1 = str(gas[1]).lstrip("b")
    gas2 = int(gas1.strip("'"))*0.01
    print(f"Today's gas price: ${gas2}/L")

    #PART 5: ADD TOTAL VALUES TO GOOGLE SHEETS
    col12Format=["Distance (km)","Total (km)","Gas Price ($)","Milage (L/100km)","Cost ($)"] #format labelling for Google Sheets
    for colCell in range(customerNumber-1): #inserts blank strings into col12Format to coincide with number of customers per transport
        col12Format.insert(1,'')
    gs.updateColumn(12, col12Format) #adds labelling format to column 12 on Google Sheets

    def travelEntries(): #prints all transport distances traveled, total distance, gas price, milage per transport, and cost to Google Sheets
        for travelCell in range(fleetSize):
            columnContents= gs.getColumn(travelCell+1) #getColumn value cannot be 0, therefore c+1
            travel=[] 
            for travelEle, value in enumerate(columnContents): #creates a list of travel values to be used for total calculation
                if type(columnContents[travelEle]) == float: #filter column content for only float datatype ()
                    travel.append(columnContents[travelEle]) 
            travelLog = travel[1:] #ignore distance element 0 as this is not a transport to customer distance
            sumTravel = 0
            for sumEle, value in enumerate(travel[1:]):
                sumTravel+=float(travel[1:][sumEle]) #calculate sum of travel distances
            travelLog.append(sumTravel) #appends sum to column list
            travelLog.append(gas2) #appends gas price to column list
            milage = random.randint(1,9)
            travelLog.append(milage) #appends milage to column list
            gasCosts = (gas2*(milage*(sumTravel/100)))*2 # *2 for traveling to customer, then back to satellite distribution center
            travelLog.append(gasCosts) #appends travel cost to column list
            for col in range(13+travelCell,14+travelCell): #writes travelLog list data to Google sheets
                gs.updateColumn(col, travelLog)
    travelEntries()
    print("Google Sheets update complete")

    #PART 6: WORD DOCUMENT REPORT GENERATION
    def reportWrite():
        if os.path.exists("c:\\users\\josh\\desktop\\python\\CourseCOMP1112\\FinalProject\\DailyTransportExpenses.docx")==False: #if the file does not exist
            doc = docx.Document()
            doc.add_heading('Daily Energy Expenses',0)
            doc.save('DailyTransportExpenses.docx') #Creates a new Word doc
        if os.path.exists("c:\\users\\josh\\desktop\\python\\CourseCOMP1112\\FinalProject\\DailyTransportExpenses.docx")==True: 
            doc = docx.Document('DailyTransportExpenses.docx') #open new Word doc
            reportTime = (datetime.datetime.now())
            doc.add_heading(f"Expenses: {reportTime}",4) #4 is a heading format variable
            rowContents = gs.getRow(transportNameRow) #A: get transport name values from Google Sheets
            namesX=[] #transport name array
            dataObj1 =[] #Word doc format array

            for nameEle, value in enumerate(rowContents): 
                if type(rowContents[nameEle]) == str and rowContents[nameEle] !='' and nameEle<11: #allows only data type str and non empty strings to be iterated
                    namesX.append(rowContents[nameEle]) #name array filled iteratively
                    dataObj1.append(f"Transport {nameEle} [{namesX[nameEle]}]") #Word doc array filled iteratively
            rowContents = gs.getRow(customerNumber+totalRowFormat) #B: get transport transport cost values from Google Sheets
            costsX=[] #cost array
            dataObj2 = [] #Word doc format array
            transportFleetCost = 0 #initial total cost value

            for costEle, value in enumerate(rowContents): #costEle is 12 when a float occurs due to formatting of the total row on the right of data on Google Sheets
                if type(rowContents[costEle]) == float and costEle>10: #allows only data type float and cells greater than column/[costEle] > 10
                    costsX.append(rowContents[costEle]) #costs array filled iteratively
                    transportFleetCost+=float(rowContents[costEle]) #total cost summed iteratively
                    dataObj2.append(f" [{costsX[costEle-totalColFormat]}]") #Word doc array filled iteratively
            for wordEle in range(fleetSize): #C: formats transport name and price in Word doc
                doc.add_paragraph(f"{dataObj1[wordEle]}:{dataObj2[wordEle]}")
            doc.add_paragraph(f"The projected total cost of running the fleet today is: ${transportFleetCost}") #adds total fleet cost/day to Word doc
        doc.save('DailyTransportExpenses.docx') #values updated 4 in a row with print(fleetSize) & print(z) & closing program & doc.save after print complete
        print("Microsoft Word update complete")
    reportWrite()

    #PART 7: AUTOMATED REPORT WRITING (GOOGLE SHEETS AND WORD) AT A SCHEDULED TIME
    time.sleep(60) #time until next report; seconds in a day = 86400

#Future additions: 
#1. User option to amend transport name, 2. User option to add or remove customer delivery
#3. Sort coordinates to minimum lengths for shortest round trip (traveling salesman problem), 4. milage is based on package mass, 5. time of transport
#6. timedInput() could be used instead of time.sleep() to give the user an option to select pause reports