import ReadExcel
import xlwt

wFile = xlwt.Workbook()

wSheet = wFile.add_sheet(ReadExcel.Sheetname)
item = ['产品代码', 'TestCase_ID', '*测试用例名', '测试目的', '优先级', '步骤', '步骤描述', '步骤描述', '预期结果']

# write item
for i in range(0, len(item)):
    wSheet.write(0, i, item[i])

# write index
crow = 0
Crow = []
print('len.All = ', len(ReadExcel.StepAll[0]))

for k in range(len(ReadExcel.StepAll[0]) + 1):
    if (k > 0):
        crow = crow + len(ReadExcel.StepAll[0][(k - 1)])
    Crow.append(crow)

for i in range(len(ReadExcel.StepAll[0])):
    for j in range(len(ReadExcel.StepAll[0][i])):
        wSheet.write((Crow[i] + j + 1), 5, str(ReadExcel.StepAll[0][i][j]))
        wSheet.write((Crow[i] + j + 1), 6, str(ReadExcel.StepAll[1][i][j]))
        wSheet.write((Crow[i] + j + 1), 7, str(ReadExcel.StepAll[2][i][j]))
    # write_merge(self, r1, r2, c1, c2) r行，c列，从0开始。其实还有格式可以选择，可研究优化
    wSheet.write_merge(Crow[i] + 1, Crow[i + 1], 0, 0, ReadExcel.ProduceCode)
    wSheet.write_merge(Crow[i] + 1, Crow[i + 1], 1, 1, ReadExcel.TcId[i])
    wSheet.write_merge(Crow[i] + 1, Crow[i + 1], 2, 2, ReadExcel.TcNm[i])
    wSheet.write_merge(Crow[i] + 1, Crow[i + 1], 3, 3, ReadExcel.TcPr[i])
    wSheet.write_merge(Crow[i] + 1, Crow[i + 1], 4, 4, ReadExcel.TC_Priority)

wFile.save(ReadExcel.TargetFileName)
print('Write done!')
