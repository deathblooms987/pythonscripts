#Python3 code to select data from Excel
import xlwings as xw
import json

#Specifying a sheet from the Excel doc
worksheet = xw.Book("Middle-East-GCC-Car-Database-by-Teoalida-SAMPLE.xlsx").sheets["Engine Specs"]

#Selecting Data from a range of cells
worksheetData = worksheet.range('C16:AE232')
worksheetDataLength = worksheetData.rows.count # 217 rows
worksheetColumnsLength = len(worksheetData.rows[0])

# print("Length: ", worksheetDataLength) # 217
# print("Row 1: ", worksheetColumnsLength)


makesData = []
make_id = 1
modelsData = []
model_id = 1
yearsData = []
vehiclesData = []
vehicle_id = 1

for index, row in enumerate(worksheetData.rows):
    makesData.append({ 'make': row[0].value, 'make_arabic': row[1].value })
    # row[0] # Make
    # row[1] # Make arabic
    modelsData.append({ 'model': row[2].value, 'model_arabic': row[3].value, 'make': row[0].value })
    # row[2] # Model
    # row[3] # Model arabic
    yearsData.append({ 'year': str(int(row[4].value)) })
    # row[4] # Year
    # row[5] # Image 1
    # row[6] # Image 2
    # row[7] # Price UAE
    # row[8] # Price KSA
    # row[9] # Country of Origin
    # row[10] # Class
    # row[11] # Body Styles
    # row[12] # Weight
    # row[13] # Good reviews
    # row[14] # Bad reviews
    # row[15] # Overview
    # row[16] # Reliability
    # row[17] # Resale value
    # row[18] # Known Problems
    # row[19] # NHTS Driver Frontal Rating
    # row[20] # Euro NCAP OVerall Adult Rating
    # row[21] # Engine Size
    # row[22] # Gearbox
    # row[23] # Power (hp)
    # row[24] # Torque
    # row[25] # Fuel Econ (L/100km)
    # row[26] # Fuel Econ (km/L)
    # row[27] # 0-100 kph (sec)
    # row[28] # Top Speed (kph)
    vehiclesData.append( {
        'make': row[0].value,
        'model': row[2].value,
        'year_id': int(row[4].value),
        'images': [row[5].value, row[6].value],
        'priceUAE': row[7].value,
        'priceKSA': row[8].value,
        'originCountry': row[9].value,
        'carClass': row[10].value,
        'bodyStyles': row[11].value,
        'weight': row[12].value,
        'reviews': [row[13].value, row[14].value],
        'overview': row[15].value,
        'reliability': row[16].value,
        'resaleValue': row[17].value,
        'knownProblems': row[18].value,
        'nhtsDriverRating': row[19].value,
        'euroNCAPRating': row[20].value,
        'engineSize': row[21].value,
        'gearbox': row[22].value,
        'horsepower': row[23].value,
        'torque': row[24].value,
        'fuelEconomy': [row[25].value, row[26].value],
        'speed': [row[27].value, row[28].value]
    } )

#Set list to get unique values
makeList = set()
yearList = set()
modelList = set()
vehicleList = set()

for item in makesData:
    if str(item) not in makeList:
        makeList.add(str(item))

for item in yearsData:
    if str(item) not in yearList:
        yearList.add(str(item))

for item in modelsData:
    if str(item) not in modelList:
        modelList.add(str(item))

for item in vehiclesData:
    if str(item) not in vehicleList:
        vehicleList.add(str(item))

makeList = list(makeList)
yearList = list(yearList)
modelList = list(modelList)
vehicleList = list(vehicleList)

#Create object to be stored in makes collection MongoDB
#class Make:
#    def __init__(self, id, name) -> None:
#        self.id = id
#        self.name = name

#Declare array size of makeList
makesArray = []
yearsArray = []
modelsArray = []
vehiclesArray = []

#Loop through and return index and name from makeList
#enumerate gets index and value from list
for index, item in enumerate(makeList):
   makesArray.append(eval(item))
   makesArray[index].update({ 'make_id': make_id })
   make_id += 1

for index, item in enumerate(yearList):
    yearsArray.append(eval(item))
    yearsArray[index].update({ 'year_id': int(yearsArray[index]['year']) })

for index, item in enumerate(modelList):
    modelsArray.append(eval(item))
    #For loop to iterate over makes to append make_id in list of models
    for makesItem in makesArray:
        if makesItem['make'] == modelsArray[index]['make']:
            modelsArray[index]['make_id'] = makesItem['make_id']
            break
    modelsArray[index].update({ 'model_id': model_id })
    del modelsArray[index]['make']
    model_id += 1

for index, item in enumerate(vehicleList):
    vehiclesArray.append(eval(item))
    for makesItem in makesArray:
        if makesItem['make'] == vehiclesArray[index]['make']:
            vehiclesArray[index]['make_id'] = makesItem['make_id']
            break
    for modelsItem in modelsArray:
        if modelsItem['model'] == vehiclesArray[index]['model']:
            vehiclesArray[index]['model_id'] = modelsItem['model_id']
            break
    del vehiclesArray[index]['make']
    del vehiclesArray[index]['model']
    vehiclesArray[index].update({'vehicle_id': vehicle_id})
    vehicle_id += 1

#Printing each object
#print("Details: ", makesArray)
with open('makes.json', 'w', encoding='utf-8') as f:
    json.dump(makesArray, f)
    f.close()

#Print for years
with open('years.json', 'w') as f:
    json.dump(yearsArray, f)
    f.close()

#print for models
with open('models.json', 'w', encoding='utf-8') as f:
    json.dump(modelsArray, f)
    f.close()

#print for vehicles
with open('vehicles.json', 'w', encoding='utf-8') as f:
    json.dump(vehiclesArray, f)
    f.close()