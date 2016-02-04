import pyodbc, csv

connect_string = 'DRIVER={SQL Server};SERVER=TASS;DATABASE=tass;UID=excel;PWD=excel'

def get_columns(tblName, cnxn):
    cursor = cnxn.cursor()
    cursor.execute("SELECT * FROM INFORMATION_SCHEMA.Columns WHERE TABLE_NAME = '%s'" %(tblName))
    return [row[3] for row in cursor]

def get_data(tblName, cnxn):
    cursor = cnxn.cursor()
    if tblName == "fim_student":
        print("trying to sort descending")
        try:
            cursor.execute("SELECT * FROM %s WHERE cmpy_code = 01 ORDER BY stud_code DESC" %(tblName))
        except:
            pass
    else:
        try:
            cursor.execute("SELECT * FROM %s WHERE cmpy_code = 01" %(tblName))
        except:
            try:
                cursor.execute("SELECT * FROM %s" %(tblName))
            except pyodbc.ProgrammingError:
                cursor.execute("SELECT * FROM [%s]" %(tblName))
    return [row for row in cursor]

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
