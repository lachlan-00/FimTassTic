import pyodbc, csv

connect_string = 'DRIVER={SQL Server};SERVER=%CHANGE%;DATABASE=%CHANGE%;UID=%CHANGE%;PWD=%CHANGE%'

def get_data(tblName, cnxn):
    cursor = cnxn.cursor()
    cursor.execute("SELECT * FROM %s" %(tblName))
    return [row for row in cursor]

def get_columns(tblName, cnxn):
    cursor = cnxn.cursor()
    cursor.execute("SELECT * FROM INFORMATION_SCHEMA.Columns WHERE TABLE_NAME = '%s'" %(tblName))
    return [row[3] for row in cursor]

def qexport(tblName):
    connection = pyodbc.connect(connect_string)
    outfile = open('%s.csv' %(tblName), 'wb')
    writer = csv.writer(outfile)
    writer.writerow(get_columns(tblName, connection))
    writer.writerows(get_data(tblName, connection))
    outfile.close()

if __name__ == "__main__":
    import sys
    qexport(sys.argv[1])
