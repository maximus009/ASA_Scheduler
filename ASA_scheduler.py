import xlrd, xlwt
import sys, os


### Eter file name 'fileName' here
fileName = sys.argv[1]
fileName = fileName.split('/')[-1]


centralBuildings = ['CAB', 'WIL', 'NAU', 'GIB', 'COC']
westBuildings = ['MON', 'MRY', 'CLK', 'PHS', 'DL1', 'DL2', 'MIN', 'CHM']

def read_excel(fileName):

    """ Reads the input excel file and returns an array of dictionaries corresponding to the worksheet """
    workBook = xlrd.open_workbook(fileName)
    workSheet = workBook.sheet_by_index(0)
    headers = [workSheet.cell(1,c).value.lower() for c in range(workSheet.ncols)]

    records = []

    for r in range(2,workSheet.nrows):
        record = {}
        for c in range(workSheet.ncols):
             record[headers[c]] = workSheet.cell(r,c).value

        records.append(record)
            
    return records


def sort_buildings(records, sort=True):
    
    """ Separates the records (and optionally sorts it by time) to central or west grounds """
    westRooms, centralRooms = [], []
    location = [record['location'] for record in roomRecords]
    for recordIndex, record in enumerate(records):
        building = record['location'].split(' ')[0]
        if building in centralBuildings:
            centralRooms.append(record)
        elif building in westBuildings:
            westRooms.append(record)
        else:
            print building,
            print "Unknown location!"

    if sort is True:
        centralRooms = sort_by_time(centralRooms)
        westRooms = sort_by_time(westRooms)

    return centralRooms, westRooms
    
def convert_time_to_int(strTime):
    hour, minute = map(int, strTime.split(' ')[0].split(':'))
    return  hour*100+minute

def sort_by_time(records):

    """ Arranges the records / ground as per the time/shifts """
    # Column End time saved as header ''
    sortedRecords = sorted(records, key=lambda rec: convert_time_to_int(rec['']))
    return sortedRecords

def write_to_book(centralRooms, westRooms, newBookName='new_'+fileName):
    
    if newBookName[-1] == 'x':
        newBookName = newBookName[:-1]

    """ Creates new workbook with separate worksheets """

    newWorkBook = xlwt.Workbook()
    sheet = newWorkBook.add_sheet('Central')
    outHeaders = ['event times', 'end time', 'location', 'event/reservation', 'organization']
    for c, header in enumerate(outHeaders):
        sheet.write(0,c, header)

    for r, record in enumerate(centralRooms):
        for c, columnName in enumerate(outHeaders):
            if columnName == 'end time':
                columnName = ''
            sheet.write(r+1,c,record[columnName])

############
    sheet = newWorkBook.add_sheet('West')
    outHeaders = ['event times', 'end time', 'location', 'event/reservation', 'organization']
    for c, header in enumerate(outHeaders):
        sheet.write(0,c, header)

    for r, record in enumerate(westRooms):
        for c, columnName in enumerate(outHeaders):
            if columnName == 'end time':
                columnName = ''
            sheet.write(r+1,c,record[columnName])
    
    print "Saving to",newBookName
    newWorkBook.save(newBookName)

if __name__=="__main__":
    roomRecords = read_excel(fileName)
    centralRooms, westRooms = sort_buildings(roomRecords)
    print len(roomRecords), len(centralRooms), len(westRooms)
    write_to_book(centralRooms, westRooms)
