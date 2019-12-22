import xlrd
import pymysql

# Open the workbook and define the worksheet
book = xlrd.open_workbook('beginner_assignment01.xlsx')

sheet = input("Enter the sheet name:")

if(book.sheet_by_name(sheet).name =='product_listing'):
	sheet = book.sheet_by_name("product_listing")
	# Establish a MySQL connection
	database = pymysql.connect (host="localhost", user = "root", password = "root", db = "python")

	# Get the cursor, which is used to traverse the database, line by line
	cursor = database.cursor()

	# Create the INSERT INTO sql query
	query = """INSERT INTO product_listing (Product_Name,Model_Name,Product_Serial_No,Group_Associated,Product_MRP) VALUES (%s, %s, %s, %s, %s)"""

	# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
	for r in range(1, sheet.nrows):
                                product		= sheet.cell(r,0).value
			model	        = sheet.cell(r,1).value
			serial		= sheet.cell(r,2).value
			group		= sheet.cell(r,3).value
			price		= sheet.cell(r,4).value
		

			# Assign values from each row
			values = (product, model, serial, group, price )

			# Execute sql Query
			cursor.execute(query, values)

	# Close the cursor
	cursor.close()

	# Commit the transaction
	database.commit()

	# Close the database connection
	database.close()

	# Print results
	print ("")
	print ("Data Imported Successfully !!")
	print ("")
	columns = str(sheet.ncols)
	rows = str(sheet.nrows)
	print ("Summary of Data imported: " + columns + " columns and " + rows + " rows to MySQL!")

elif(book.sheet_by_name(sheet)):
	sheet = book.sheet_by_name("group_listing")
	database = pymysql.connect (host="localhost", user = "root", passwd = "root", db = "python")

	# Get the cursor, which is used to traverse the database, line by line
	cursor = database.cursor()

	# Create the INSERT INTO sql query
	query = """INSERT INTO group_listing(group_name,group_description,isActive) VALUES (%s, %s, %s)"""

	# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
	for r in range(1, sheet.nrows):
			group_name          = sheet.cell(r,0).value
			group_description   = sheet.cell(r,1).value
			isActive            = sheet.cell(r,2).value
		

			# Assign values from each row
			values = (group_name,group_description,isActive )

			# Execute sql Query
			cursor.execute(query, values)

	# Close the cursor
	cursor.close()

	# Commit the transaction
	database.commit()

	# Close the database connection
	database.close()

	# Print results
	print ("")
	print ("Data Imported Successfully !!")
	print ("")
	columns = str(sheet.ncols)
	rows = str(sheet.nrows)
	print ("Summary of Data imported: " + columns + " columns and " + rows + " rows to MySQL!")






