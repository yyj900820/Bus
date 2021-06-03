import time
import xlrd
import xlwt
start_time = time.time()
no_sheet = 0
read_file_dir = "poi-320509-150700-04131106-1.xls"  # poi文件
myWordbookr = xlrd.open_workbook(read_file_dir)
mySheetsr = myWordbookr.sheets()
mySheetr = mySheetsr[no_sheet]
nrows = mySheetr.nrows
busline=[]
name = "公交线路" + time.strftime("%m%d%H%M", time.localtime())
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet("1", cell_overwrite_ok=True)
sheet.write(0, 0, 'busline')
for i in range(1, nrows):
    cityname = mySheetr.cell_value(i, 5)
    bus = mySheetr.cell_value(i, 3)
    print("bus", bus)
    for j in range(len(bus.split(';'))):
        bus1 = bus.split(';')[j]
        busline.append(bus1)


print("bl",busline)
for k in range(len(busline)):
    sheet.write(k+1, 0, busline[k])
end_time = time.time()
book.save(r'' + name + '.xls')
print('写入成功,总用时%.2f秒' % (end_time - start_time))