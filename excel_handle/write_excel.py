import xlrd
import xlwt

file = '仲景中药配方颗粒柘城中医院采购明细表.xls'
file2 = '仲景中药配方颗粒南阳市第二人民医院价格.xls'
file3 = '仲景中药配方颗粒南阳张仲景医院价格.xls'
file4 = '仲景中药配方颗粒永城价格.xlsx'


def get_data():
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet_zhecheng = wb.sheet_by_index(0)  # 通过索引获取表格
    medicine_name_zhecheng = sheet_zhecheng.col_values(0)
    medicine_name_zhecheng = medicine_name_zhecheng[2:len(medicine_name_zhecheng) - 1]
    medicine_price_zhecheng = sheet_zhecheng.col_values(5)
    medicine_price_zhecheng = medicine_price_zhecheng[2:len(medicine_price_zhecheng) - 1]

    # 南阳第二
    wb1 = xlrd.open_workbook(filename=file2)  # 打开文件
    sheet_nanyanger = wb1.sheet_by_index(0)
    medicine_name_nayanger = sheet_nanyanger.col_values(1)
    medicine_price_nayanger = sheet_nanyanger.col_values(6)
    # 南阳第二
    nanyang_medicine_name_list = []
    nanyang_medicine_price_list = []
    for medicine_name in medicine_name_zhecheng:
        if medicine_name in medicine_name_nayanger:
            index = medicine_name_nayanger.index(medicine_name)
            nanyang_medicine_name_list.append(medicine_name)
            nanyang_medicine_price_list.append(medicine_price_nayanger[index])
        else:
            nanyang_medicine_name_list.append('无')
            nanyang_medicine_price_list.append(0)

    # 南阳仲景
    wb2 = xlrd.open_workbook(filename=file2)  # 打开文件
    sheet_nanyang_zhongjing = wb2.sheet_by_index(0)
    medicine_name_nanyang_zhongjing = sheet_nanyang_zhongjing.col_values(1)
    medicine_price_nanyang_zhongjing = sheet_nanyang_zhongjing.col_values(6)

    nanyang_zhongjing_medicine_name_list = []
    nanyang_zhongjing_medicine_price_list = []

    for medicine_name in medicine_name_zhecheng:
        if medicine_name in medicine_name_nanyang_zhongjing:
            index = medicine_name_nanyang_zhongjing.index(medicine_name)
            nanyang_zhongjing_medicine_name_list.append(medicine_name)
            nanyang_zhongjing_medicine_price_list.append(medicine_price_nanyang_zhongjing[index])
        else:
            nanyang_zhongjing_medicine_name_list.append('无')
            nanyang_zhongjing_medicine_price_list.append(0)


    # 永城
    wb3 = xlrd.open_workbook(filename=file4)  # 打开文件
    sheet_yongcheng = wb3.sheet_by_index(0)
    medicine_name_yongcheng = sheet_yongcheng.col_values(1)
    medicine_price_yoncheng = sheet_yongcheng.col_values(6)

    # 永城
    yongcheng_medicine_name_list = []
    yongcheng_medicine_price_list = []
    for medicine_name in medicine_name_zhecheng:
        if medicine_name in medicine_name_yongcheng:
            index = medicine_name_yongcheng.index(medicine_name)
            yongcheng_medicine_name_list.append(medicine_name)
            yongcheng_medicine_price_list.append(medicine_price_yoncheng[index])
        else:
            yongcheng_medicine_name_list.append('无')
            yongcheng_medicine_price_list.append(0)

    return medicine_name_zhecheng, medicine_price_zhecheng, nanyang_medicine_price_list, \
           nanyang_zhongjing_medicine_price_list, yongcheng_medicine_price_list


# 设置表格样式
def set_style(name, height, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    # font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style


# 写Excel
def write_excel():
    medicine_name_zhecheng, medicine_price_zhecheng, nanyang_medicine_price_list, nanyang_zhongjing_medicine_price_list,yongcheng_medicine_price_list = get_data()
    f = xlwt.Workbook()
    sheet1 = f.add_sheet('药品价格', cell_overwrite_ok=True)
    sheet1.col(0).width = 256 * 30
    sheet1.col(1).width = 256 * 20
    sheet1.col(2).width = 256 * 30
    sheet1.col(3).width = 256 * 20
    sheet1.col(4).width = 256 * 20

    row0 = ["药品名字", "柘城报价", "南阳第二人民医院报价", "南阳张仲景医院报价", "仲景永城药品报价"]
    colum0 = medicine_name_zhecheng
    colum1 = medicine_price_zhecheng
    colum2 = nanyang_medicine_price_list
    colum3 = nanyang_zhongjing_medicine_price_list
    colum4 = yongcheng_medicine_price_list


    # colum0 = ["张三","李四","恋习Python","小明","小红","无名"]
    # 写第一行 head
    for i in range(0, len(row0)):
        print(i)
        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    # 写第一列 药品名字
    for i in range(0, len(colum0)):
        sheet1.write(i + 1, 0, colum0[i], set_style('Times New Roman', 220, True))
    # 写第二列 药品价格柘城
    for i in range(0, len(colum1)):
        sheet1.write(i + 1, 1, colum1[i], set_style('Times New Roman', 220, True))
    # 写第三列 药品价格南阳第二人民医院
    for i in range(0, len(colum2)):
        sheet1.write(i + 1, 2, colum2[i], set_style('Times New Roman', 220, True))
    # 写第三列 药品价格南阳仲景医院
    for i in range(0, len(colum3)):
        sheet1.write(i + 1, 3, colum3[i], set_style('Times New Roman', 220, True))
    # 写第三列 永城
    for i in range(0, len(colum4)):
        sheet1.write(i + 1, 4, colum4[i], set_style('Times New Roman', 220, True))
    f.save('药品价格对比.xls')


if __name__ == '__main__':
    write_excel()
# get_data()
