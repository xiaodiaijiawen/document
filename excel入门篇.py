import xlwings as xw
app=xw.App(visible=True,add_book=False)
app.display_alerts=True
app.screen_updating=True
#文件位置：filepath，打开test文档，然后保存，关闭，结束程序
filepath=r'C:\Users\WUD\Desktop\导入.xls'
#filepath=r'C:\Users\WUD\Desktop\V90更新文档目录_20170815.xlsx'
wb=app.books.open(filepath)
#print(wb)
sht=wb.sheets.active
#sht=wb.sheets['Tabelle1']
#print(sht)

a=sht.range('A4').expand().value
print(a)
wb.close()





#app=xw.App(visible=True,add_book=False)
#app.display_alerts=True
#app.screen_updating=True
#filepath=r'C:\Users\WUD\Desktop\微信数据记录模板(1) - 副本.xlsx'
filepath=r'C:\Users\WUD\Desktop\微信数据记录模板.xlsx'
wb=app.books.open(filepath)
#sht=wb.sheets.active
sht=wb.sheets['用户数据源']
 # wb就是新建的工作簿(workbook)，下面则对wb的sheet1的A1单元格赋值
i=1

#print(sht.range(i,3).value)
while sht.range(i,3).value!=None:
    i=i+1
sht.range(i,3).value=a
m=32


wb.save(r'C:\Users\WUD\Desktop\321.xlsx')
wb.save(r'C:\Users\WUD\Desktop\微信数据记录模板(1) - 副本.xlsx')
wb.close()
app.quit()