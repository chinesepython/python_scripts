import xlrd


file = '仲景中药配方颗粒柘城中医院采购明细表.xls'
file2 = '仲景中药配方颗粒南阳市第二人民医院价格.xls'
file3 = '仲景中药配方颗粒南阳张仲景医院价格.xls'
file4 = '仲景中药配方颗粒永城价格.xlsx'

def read_excel():
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet_zhecheng = wb.sheet_by_index(0)  # 通过索引获取表格
    medicine_name_zhecheng1 = sheet_zhecheng.col_values(0)
    medicine_name_zhecheng = medicine_name_zhecheng1[2:len(medicine_name_zhecheng1)-1]
    medicine_price_zhecheng1 = sheet_zhecheng.col_values(5)
    medicine_price_zhecheng = medicine_price_zhecheng1[2:len(medicine_price_zhecheng1)-1]

    # 南阳第二人民医院
    wb1 = xlrd.open_workbook(filename=file2)  # 打开文件
    sheet_nanyanger = wb1.sheet_by_index(0)
    medicine_name_nayanger = sheet_nanyanger.col_values(1)
    medicine_price_nayanger = sheet_nanyanger.col_values(6)

    # 永城
    wb2 = xlrd.open_workbook(filename=file4)  # 打开文件
    sheet_yongcheng = wb2.sheet_by_index(0)
    medicine_name_yongcheng = sheet_yongcheng.col_values(1)
    medicine_price_yoncheng = sheet_yongcheng.col_values(6)



    nanyang_medicine_name_list = []
    nanyang_medicine_price_list = []
    for medicine_name in medicine_name_zhecheng:
        if medicine_name in medicine_name_nayanger:
            index = medicine_name_nayanger.index(medicine_name)
            nanyang_medicine_name_list.append(medicine_name)
            nanyang_medicine_price_list.append(medicine_price_nayanger[index])
        else:
            nanyang_medicine_name_list.append('无')
            nanyang_medicine_price_list.append('0')
    # print(medicine_name_zhecheng)
    # print(nanyang_medicine_name_list)
    # print(nanyang_medicine_price_list)
    #
    # print(len(medicine_name_zhecheng))
    # print(len(nanyang_medicine_name_list))
    # print(len(nanyang_medicine_price_list))

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

    print(len(yongcheng_medicine_price_list))
    print(yongcheng_medicine_price_list)
    print(len(medicine_name_zhecheng))
    print(yongcheng_medicine_name_list)
    return medicine_name_zhecheng


if __name__ == '__main__':
    read_excel()
