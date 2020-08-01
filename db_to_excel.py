import psycopg2
import xlsxwriter

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('info_from_db.xlsx')
worksheet = workbook.add_worksheet()

# Set up some formats to use.
red = workbook.add_format({'color': 'red'})
blue = workbook.add_format({'color': 'blue'})
cell_format = workbook.add_format({'align': 'center',
                                   'valign': 'vcenter',
                                   'border': 1})

space_format = workbook.add_format({'align': 'center',
                                   'valign': 'vcenter'})

# We can only write simple types to merged ranges so we write a blank string.
worksheet.merge_range('D2:E2', "", cell_format)

# Create Table Headings
worksheet.write('B2', 'ID', cell_format)
worksheet.write_rich_string('C2',
                            'The ',
                            blue, 'Model',
                            cell_format)
worksheet.write_rich_string('D2',
                            'Phone ',
                            red, 'price',
                            cell_format)


# function to write one row
def write_data(worksheet_, row_id, id_, model, price):
    worksheet_.write('B' + str(row_id), id_)
    worksheet_.write('C' + str(row_id), model)

    worksheet.merge_range(f'D{row_id}:E{row_id}', "", space_format)
    worksheet_.write_rich_string('D' + str(row_id),
                                 ' ',
                                 red, str(price),
                                 space_format)


try:
    connection = psycopg2.connect(user="mdashwmj",
                                  password="Yt0riCoIAke8rUFNHS4sDuKw6NKBW5GU",
                                  host="ruby.db.elephantsql.com",
                                  port="5432",
                                  database="mdashwmj")

    cursor = connection.cursor()

    postgreSQL_select_Query = "select * from mobile"

    cursor.execute(postgreSQL_select_Query)
    print("Selecting rows from mobile table using cursor.fetchall")
    mobile_records = cursor.fetchall()

    print("Print each row and it's columns values")
    for idx, row in enumerate(mobile_records):
        print("Id = ", row[0], )
        print("Model = ", row[1])
        print("Price  = ", row[2], "\n")
        write_data(worksheet, idx + 3, row[0], row[1], row[2])  # idx + 3 - because we start to write data from 3rd cell

except psycopg2.DatabaseError as error:
    print("Error while working with PostgreSQL:", error)
finally:
    # closing database connection.
    if connection:
        cursor.close()
        connection.close()
        print("PostgreSQL connection is closed")

workbook.close()
