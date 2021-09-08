import os
import sys
import uuid
import zipfile
import argparse
import xlsxwriter


temp_path = os.path.dirname(os.path.abspath(__file__)) + os.path.sep + "tmp"
if not os.path.exists(temp_path):
    os.makedirs(temp_path)

my_parser = argparse.ArgumentParser(description='Coverts csv to xlsx and viceversa')

my_parser.add_argument('--csv', metavar='csv', type=str, help='CSV file path')
my_parser.add_argument('--csv_delimiter', metavar='csv_delimiter', type=str, help='CSV file delimiter')
# my_parser.add_argument('xlsx', metavar='xlsx', type=str, help='xlsx file path')
# my_parser.add_argument('out', metavar='out', type=str, help='output file path')

SHEET_LIMIT = 50000
MAX_SHEETS_PER_XLS = 7
MAX_LINES_IN_ONE_SHOT = 300000

def zipper(zip_file_name, files):
    zip_file_name = '{}{}{}.zip'.format(temp_path, os.path.sep, zip_file_name)
    print("Zipping {} xlsx files to {}".format(len(files), zip_file_name))
    with zipfile.ZipFile(zip_file_name,'w') as zip_:
        for file_ in files:
            zip_.write(temp_path + os.path.sep + file_)
    print("file:{}".format(zip_file_name))
    return os.path.basename(zip_file_name)
        

def create_xlsx(headers, data=None, data_list=[], local=False, sheet_limit=SHEET_LIMIT):
    main_file_name = str(uuid.uuid4())+".xlsx"
    sheet_number = 1
    workbook = xlsxwriter.Workbook(temp_path + os.path.sep + main_file_name)
    worksheet = workbook.add_worksheet(name="Sheet{}".format(sheet_number))
    row = 0
    col = 0
    total_files = [main_file_name]
    for header in headers:
        worksheet.write(row, col, header)
        col += 1

    if not local:
        for col_data in data:
            if row != 0:
                print("progress({}/100)".format(int(sheet_number*(row/len(data)*100))), end="\r")
            row += 1
            col = 0
            for header in headers:
                #print("header--> {}:data-->{}:type--->{}".format(header,col_data[header],type(col_data[header])))
                worksheet.write(row, col, col_data[header])
                col += 1

            if sheet_number < MAX_SHEETS_PER_XLS:
                if row >= sheet_limit:
                    row = 0
                    col = 0
                    sheet_number += 1
                    worksheet = workbook.add_worksheet(name="Sheet{}".format(sheet_number))
                    for header in headers:
                        worksheet.write(row, col, header)
                        col += 1
            else:
                if row < sheet_limit:
                    continue
                workbook.close() 
                file_name = "{}_part-{}.xlsx".format(os.path.splitext(main_file_name)[0], len(total_files))
                sheet_number = 1
                workbook = xlsxwriter.Workbook(temp_path + os.path.sep + file_name)
                worksheet = workbook.add_worksheet(name="Sheet{}".format(sheet_number))
                row = 0
                col = 0
                for header in headers:
                    worksheet.write(row, col, header)
                    col += 1
                total_files.append(file_name)
                
    else:
        for data in data_list:
            for col_data in data:
                row += 1
                col = 0
                for header in headers:
                    #print("header--> {}:data-->{}:type--->{}".format(header,col_data[header],type(col_data[header])))
                    worksheet.write(row, col, col_data[header])
                    col += 1

                if sheet_number < MAX_SHEETS_PER_XLS:
                    if row >= sheet_limit:
                        row = 0
                        col = 0
                        sheet_number += 1
                        worksheet = workbook.add_worksheet(name="Sheet{}".format(sheet_number))
                        for header in headers:
                            worksheet.write(row, col, header)
                            col += 1
                else:
                    if row < sheet_limit:
                        continue
                    workbook.close() 
                    file_name = "{}_part-{}.xlsx".format(os.path.splitext(main_file_name)[0], len(total_files))
                    sheet_number = 1
                    workbook = xlsxwriter.Workbook(temp_path + os.path.sep + file_name)
                    worksheet = workbook.add_worksheet(name="Sheet{}".format(sheet_number))
                    row = 0
                    col = 0
                    for header in headers:
                        worksheet.write(row, col, header)
                        col += 1
                    total_files.append(file_name)  

    workbook.close()      

    if len(total_files) == 1:
        return main_file_name
    else:
        return zipper(os.path.splitext(main_file_name)[0], total_files)

def create_data(headers, row_data):
    info = {}
    for i in range(0, len(headers)):
        info[headers[i]] = row_data[i]
    return info

def convert_csv_to_xlsx(csv_file, xlsx_file, delimiter=","):
    with open(csv_file, encoding="utf-8") as f:
        data_list = []
        headers = f.readline().rstrip().replace('"',"").split(delimiter)
        missed_lines_count = 0
        print(headers)
        lines_loaded = 0
        final_files = []
        while(line:=f.readline()):
            lines_loaded += 1
            print("lines-loaded----{}".format(lines_loaded),end="\r")
            line = line.rstrip()
            try:
                data_list.append(create_data(headers, line.replace('"',"").split(delimiter)))
            except:
                print(line)
                missed_lines_count += 1
            if len(data_list) >= MAX_LINES_IN_ONE_SHOT:
                print("Total lines missed--->{}".format(missed_lines_count))    
                final_file = create_xlsx(headers, data=data_list)
                print("Your xlsx path is {}".format( final_file))
                final_files.append(final_file)
                data_list = []
        print("Total lines missed--->{}".format(missed_lines_count))    
        final_files.append(create_xlsx(headers, data=data_list))
        print("Your xlsx files are {}".format( ",".join(final_files)))

if __name__ == "__main__":
    args = my_parser.parse_args()
    if not args.csv:
        print("Invalid info provided For Help: python covert.py --help")
        sys.exit(0)


    if args.csv:
        out_file = os.path.splitext(args.csv)[0] + ".xlsx"
        delimiter = args.csv_delimiter if args.csv_delimiter else None
        if delimiter:
            convert_csv_to_xlsx(args.csv, out_file, delimiter)
