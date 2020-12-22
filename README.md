# hellowword
lean github

## movie Excel sheet from multiple workbook to one workbook

import xlwings as xw
app = xw.App(visible=False)
target_wb = xw.Book()
target_api = target_wb.sheets[0].api
for f in glob.glob('myobreport\\*.xlsx'):
    xw.Book(f).sheets['Sheet1'].api.Copy(Before = target_api)
    xw.Book(f).close()


target_wb.save('TargetExcel.xlsx')
app.kill()
