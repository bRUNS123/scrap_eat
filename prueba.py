import xlsxwriter

k = 0 
while k < 1000:    
     workbook = xlsxwriter.Workbook(f'./{k}.xlsx')
     worksheet = workbook.add_worksheet()
     workbook.close()

     k = k+1
