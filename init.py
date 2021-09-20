#connection=mysql.connector.connect(user='root', password='',host='localhost')

#mycursor.execute("CREATE DATABASE yokmak DEFAULT CHARACTER SET 'utf8'")


from getpass import getpass
from mysql.connector import connect, Error
import xlrd
import zipfile

def connection():
	try:
		with connect(
        	host="localhost",
        	user=input("root"),
        	password=getpass(""),
    	) as connection:
        		create_db_query = "CREATE DATABASE yokmak"
        		with connection.cursor() as cursor:
        			cursor.execute(create_db_query)
	except Error as e:
    		print(e)


	create_principaltable = """
		CREATE TABLE principal_data(
		id INT(50) NOT NULL AUTO_INCREMENT PRIMARY KEY,
		panel_number VARCHAR(20) NOT NULL,
		job_number VARCHAR(20) NOT NULL,
		job_name VARCHAR(50) NOT NULL,
		seal BOOLEAN NOT NULL,
		modbus_id INT(20) NOT NULL,
		id_second NOT NULL FOREIGN KEY(id_second)
		)ENGINE=InnoDB
	"""

	create_secondtable = """
	CREATE TABLE second_data(
		id INT(30) NOT NULL AUTO_INCREMENT PRIMARY KEY,
		id_FK NOT NULL FOREIGN KEY(id_second),
		serial_number VARCHAR(30) NOT NULL,
		meter_no INT(10) NOT NULL
		)ENGINE=InnoDB
	"""

	with connection.cursor() as cursor:
		cursor.execute(create_principaltable)
		cursor.execute(create_secondtable)
		connection.commit()

def decompress():
	dirarchive=os.getcwd()
	with zipfile.ZipFile('Yok-Mak.zip', 'r') as zip_ref: #Se debe colocar el nombre del archivo .zip
		zip_ref.extractall(dirarchive+'newdir') #Aquí abajo se coloca la ruta del archivo .zip

def extract():
	#Panel number(2,3), Job number(3,3), Seal(2,9), type(27,1), 
	#modbusid(32,2), serialnumber(49...,2) y meterno(49...,1)
	dirarchive=os.getcwd()+'/newdir'
	listarchives=os.listdir(dirarchive)

	for archives in listarchives:
		actual=xlrd.open_workbook(dirarchive+archives)
		sheet = wb.sheet_by_index(0)
		#loc = ("D:/Cursos/Yok-Mak/-32DPEA.xls")
		#wb = xlrd.open_workbook(loc) 
		panel_numberextract=sheet.cell_value(2, 3)
		job_number=sheet.cell_value(3, 3)
		if len(sheet.cell_value(3, 3)) == 0:
			seal=0
		else:
			seal=1
		taipe=sheet.cell_value(27,1)
		modbusid=sheet.cell_value(32,2)
		i=49
		fk_query="SELECT id FROM principal_data where panel_number=panel_numberextract"
		with connection.cursor() as cursor:
			result = cursor.execute(fk_query)
		while(len(sheet.cell_value(i,2))!=0):
			serialnumber=sheet.cell_value(i,2)
			meterno=sheet.cell_value(i,1)
			insertSN = """
			INSERT INTO second_data (id_FK,serial_number,meter_no) VALUES
				(result,serialnumber,meterno)
			"""
			i=i+1
			with connection.cursor() as cursor:
				cursor.execute(insert_movies_query)
				connection.commit()


if __name__ == '__main__':		
	connection()
	decompress()
	extract()
