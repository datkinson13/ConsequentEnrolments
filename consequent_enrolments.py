from xlrd import open_workbook
import datetime
import xlwt
import tkinter
import tkinter.filedialog

tkinter.Tk().withdraw()
file = tkinter.filedialog.askopenfilename()

workbook = open_workbook(file)

class EnrolmentRecord:
    def __init__(self, username, full_name, catalogue_item, created_date, enrolment_status, completed_date, expiry_date):
        self.username = username
        self.full_name = full_name
        self.catalogue_item = catalogue_item
        self.created_date = created_date
        self.enrolment_status = enrolment_status
        self.completed_date = completed_date
        self.expiry_date = expiry_date
    
    def __str__(self):
        record = "Username: {0}\nFull Name: {1}\nCatalogue Item: {2}\nCreated Date: {3}\nEnrolment Status: {4}\nCompleted Date: {5}\nExpiry Date: {6}\n".format(self.username, self.full_name, self.catalogue_item, self.created_date, self.enrolment_status, self.completed_date, self.expiry_date)
        return record

def read_records(workbook):
    for sheet in workbook.sheets():
        items = []
        rows = []

        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

        for row in range(1, number_of_rows):
            values = []

            for col in range(number_of_columns):
                value = (sheet.cell(row, col).value)
                
                try:
                    value = str(int(value))
                except ValueError:
                    pass
                finally:
                    values.append(value)
            
            item = EnrolmentRecord(*values)
            items.append(item)

    return items

def write_records(records):
    counter = 2

    upload_workbook = xlwt.Workbook()
    upload_workbook_sheet = upload_workbook.add_sheet('Data')

    # date_format = xlwt.XFStyle()
    # date_format.num_format_str = 'dd/mm/yyyy'
    date_format = xlwt.easyxf(num_format_str='dd/mm/yyyy')

    upload_workbook_sheet.write(1, 0, "OwnerId")
    upload_workbook_sheet.write(1, 1, "User Id")
    upload_workbook_sheet.write(1, 2, "Lesson Status")
    upload_workbook_sheet.write(1, 3, "Date Completed")
    upload_workbook_sheet.write(1, 4, "Date Created")
    upload_workbook_sheet.write(1, 5, "Expired Date")
    upload_workbook_sheet.write(1, 6, "Comments")

    for record in records:
        upload_workbook_sheet.write(counter, 0, "ARU_LCI_100")
        upload_workbook_sheet.write(counter, 1, record.username)
        upload_workbook_sheet.write(counter, 2, record.enrolment_status)
        upload_workbook_sheet.write(counter, 3, datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(record.completed_date) - 2), date_format) # date is wrong, needs work
        upload_workbook_sheet.write(counter, 4, datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(record.completed_date) - 2), date_format)
        # upload_workbook_sheet.write(counter, 5, datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + int(record.expiry_date) - 2), date_format)
        upload_workbook_sheet.write(counter, 6, "Imported record - Consequent Smart Rugby (LCI) from relevant courses - " + datetime.date.today().strftime('%Y%m%d'))

        counter += 1
    
    upload_workbook.save('test.xls')

records = read_records(workbook)
write_records(records)