import os
from openpyxl import load_workbook



def handle_dhl(fullpath, select_filename, tipps_value, root, progressbar):
    # openpyxl 索引从1开始
    col_list = []
    col_list1 = [('Invoice Nr', 4, 1),
                 ('Invoice Date', 3, 2),
                 ('Pu Date', 12, 3),
                 ('Identcode', 13, 4),
                 ('货主代码', 14, 5),
                 ('发货渠道', 15, 6),
                 ('目的地国家', 16, 7),
                 ('K5打单金额', 17, 8),
                 ('件数', 18, 9),
                 ('Shippers Reference', 19, 10),
                 ('Pcs', 20, 11),
                 ('Wgt', 22, 12),
                 ('Wgt_Abr', 23, 13)]
    col_list2 = []  # 费用类型
    col_list3_part = [('Total', 27)]
    col_list3 = []
    col_fee_type = 11  # 原表费用类型位置，index从1开始
    col_cost = 27  # 附加费

    if os.path.splitext(select_filename)[1] == '.xlsm':
        wb = load_workbook(fullpath, read_only=False, keep_vba=True)  # 坑！！！xlsm文件True，xlsx文件False
    else:
        wb = load_workbook(fullpath, read_only=False, keep_vba=False)

    tipps_value.set('提示:文件正在处理...')
    root.update()

    ws = wb[wb.sheetnames[0]]
    maxrow = ws.max_row

    # 获取费用类型数据
    tuple_fee_type = ws['K']
    list_fee_type = []
    for i in range(len(tuple_fee_type)):
        list_fee_type.append(list(tuple_fee_type)[i].value)

    # 去重
    list_fee_type = list(set(list_fee_type))
    list_fee_type.remove("Prod")
    list_fee_type = ["Charge Amount", "Extra Charge Amount"]


    # 将费用类型赋值给col_list2
    m = 1
    for i in range(len(list_fee_type)):
        if list_fee_type[i] == '费用描述':
            continue
        temp_list = [list_fee_type[i], col_fee_type, len(col_list1) + m]
        col_list2.append(tuple(temp_list))
        m = m + 1
    # 添加col_list3在新表中的列索引
    for i in range(len(col_list3_part)):
        temp_list = list(col_list3_part[i])
        temp_list.append(len(col_list1) + len(col_list2) + i + 1)
        col_list3.append(tuple(temp_list))

    # print(col_list3)

    # 合并col_list
    col_list.extend(col_list1)
    col_list.extend(col_list2)
    col_list.extend(col_list3)

    # 新建sheet
    try:
        wb.get_sheet_by_name('已处理')
        wb.remove(wb['已处理'])
        # print('删除之前生成的sheet')
    except KeyError:
        pass

    # print('新建sheet')
    wb.create_sheet('已处理')
    new_ws = wb.get_sheet_by_name('已处理')

    # 判断是否是对应的文件
    # print(ws['M1'].value)
    if ws['M1'].value != 'Identcode':
        tipps_value.set('文件格式不正确')
        root.update()
        return 'file false'

    for i in range(len(col_list)):
        new_ws.cell(1, col_list[i][2]).value = col_list[i][0]
    # 按行读取，第一行是标题行
    for i in range(2, maxrow+1):
        # print("i="+str(i))
        #  更新界面processbar
        progressbar['value'] = (i/maxrow)*100
        root.update()

        new_max_row = new_ws.max_row

        if ws.cell(i, col_list[3][1]).value == ws.cell(i - 1, col_list[3][1]).value:
            # 如果Identcode单号相同
            if ws.cell(i, col_list[3][1]).value is not None:

                for fee in col_list2:
                    #  附加费用遍历
                    if fee[0] == "Charge Amount":
                        # print(ws.cell(i, fee[1]).value)
                        if len(str(ws.cell(i, fee[1]).value)) == 9:

                            if new_ws.cell(new_max_row, fee[2]).value is None:
                                new_ws.cell(new_max_row, fee[2]).value = 0
                            if ws.cell(i, col_cost).value is None:
                                ws.cell(i, col_cost).value = 0

                            try:
                                new_ws.cell(new_max_row, fee[2]).value = float(new_ws.cell(new_max_row, fee[2]).value) + float(ws.cell(i, col_cost).value)
                            except ValueError:
                                pass
                                # print(new_max_row)
                                # print(fee[2])
                                # print(new_ws.cell(new_max_row, fee[2]).value)
                                # print(ws.cell(i, col_cost).value)
                    if fee[0] == "Extra Charge Amount":
                        if len(str(ws.cell(i, fee[1]).value)) < 9:
                            if new_ws.cell(new_max_row, fee[2]).value is None:
                                new_ws.cell(new_max_row, fee[2]).value = 0
                            if ws.cell(i, col_cost).value is None:
                                ws.cell(i, col_cost).value = 0

                            try:
                                new_ws.cell(new_max_row, fee[2]).value = float(new_ws.cell(new_max_row, fee[2]).value) + float(ws.cell(i, col_cost).value)
                            except ValueError:
                                pass
                                # print(new_max_row)
                                # print(fee[2])
                                # print(new_ws.cell(new_max_row, fee[2]).value)
                                # print(ws.cell(i, col_cost).value)
                # print("total:"+str(col_list3[0][1]))

                if ws.cell(i, col_list3[0][1]).value != 0:
                    # 主单号总金额

                    new_ws.cell(new_max_row, col_list3[0][2]).value = float(new_ws.cell(new_max_row, col_list3[0][2]).value) + float(ws.cell(i, col_list3[0][1]).value)
                    # print(new_ws.cell(new_max_row, col_list1[5][2]).value, " is ", float(ws.cell(i, col_list3[1][1]).value))
            else:
                for x in range(len(col_list1)):
                    new_ws.cell(new_max_row + 1, col_list1[x][2]).value = ws.cell(i, col_list1[x][1]).value
                for fee in col_list2:
                    if fee[0] == "Charge Amount":
                        if len(str(ws.cell(i, fee[1]).value)) == 9:
                            new_ws.cell(new_max_row + 1, fee[2]).value = ws.cell(i, col_cost).value
                    if fee[0] == "Extra Charge Amount":
                        if len(str(ws.cell(i, fee[1]).value)) < 9:
                            new_ws.cell(new_max_row + 1, fee[2]).value = ws.cell(i, col_cost).value
                for x in range(len(col_list3)):
                    new_ws.cell(new_max_row + 1, col_list3[x][2]).value = ws.cell(i, col_list3[x][1]).value
                # print(ws.cell(i, 53).value)
        else:
            for x in range(len(col_list1)):
                new_ws.cell(new_max_row + 1, col_list1[x][2]).value = ws.cell(i, col_list1[x][1]).value
            for fee in col_list2:
                if fee[0] == "Charge Amount":
                    if len(str(ws.cell(i, fee[1]).value)) == 9:
                        new_ws.cell(new_max_row + 1, fee[2]).value = ws.cell(i, col_cost).value
                if fee[0] == "Extra Charge Amount":
                    if len(str(ws.cell(i, fee[1]).value)) < 9:
                        new_ws.cell(new_max_row + 1, fee[2]).value = ws.cell(i, col_cost).value
            for x in range(len(col_list3)):
                new_ws.cell(new_max_row + 1, col_list3[x][2]).value = ws.cell(i, col_list3[x][1]).value
            # print(ws.cell(i, 53).value)
    try:
        wb.save(fullpath)
        wb.close()
    except PermissionError:
        tipps_value.set('文件已被占用,请先关闭')
        root.update()
        return 'please close file'

    return 'finish'
