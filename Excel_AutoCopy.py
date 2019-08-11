import re
import xlrd
import xlwt
#from xlutils.copy import copy as xl_copy

global file_name

sheet_name = []


# Get file name or sheet name

def get_name(name_str):


    string = name_str.split('=')
    name = string[1].strip()
                                         
    print(name)
    return name



    

def main():

    # Open Config file
    config_file = open("Config.txt")

    while True:

        
        content = config_file.readline()

        if len(content) != 0:
            if re.match("file_name",content):
                file_name = get_name(content)
                print(file_name)

            if re.match("sheet\d+",content):
                sheet_name.append(get_name(content))
                
        else:
            break

    wb = xlwt.Workbook(file_name)
    
    # Setting excel format style.
    style = xlwt.XFStyle()
    border_line = xlwt.Borders()
    border_line.top = 1
    border_line.left = 1
    border_line.right= 1
    border_line.bottom = 1

    style.borders = border_line

    for i in range(len(sheet_name)):

        # Open the file that will be read
        print(i)
        book = xlrd.open_workbook(sheet_name[i]+ '.xlsx')
        

        # Read sheet1
        sh = book.sheet_by_index(0)

        ws = wb.add_sheet(sheet_name[i])

        for rx in range(sh.nrows):
            for cl in range(sh.ncols):
                ws.write(rx,cl,sh.cell_value(rx,cl),style)


        i += 1

    
    wb.save(file_name)
    print(sheet_name)
    config_file.close()





if __name__ == "__main__":
    main()                

            


