import xlsxwriter


workbook = xlsxwriter.Workbook("C:\\Users\\Gh0sT\\Desktop\\WORKBOOK\\LESJOFORS.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write(0 , 0, "daaaaaa")
workbook.close()