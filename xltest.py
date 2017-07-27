# coding:utf-8
import pypyodbc
import xlsxwriter
import arcpy as GP
import os
# import xlrd
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


# a test for git
# location = raw_input("输入mdb数据库文件，如F:/learn/2014p.mdb ")
# out_path = raw_input("输入保存文件路径，如F:/learn ")
location = GP.GetParameterAsText(0)
out_path = os.path.dirname(location)
conn = pypyodbc.connect('Driver={Microsoft Access Driver (*.mdb)};DBQ=' + location)
cur = conn.cursor()
data = cur.execute("select left(zldwdm,12),dlbm,dlmj from t_hzmj")
records = data.fetchall()
workbook = xlsxwriter.Workbook(out_path + "/现状统计.xlsx")
worksheet = workbook.add_worksheet("现状地类统计")


def write_dlmc_row():
    classify = workbook.add_format({'bold': 1, 'text_h_align': 2, 'text_v_align': 2})
    title = workbook.add_format({'bold': 1, 'align': 'center', 'font': 16})

    row_mc = ['小计', '水田', '水浇', '旱地',
              '小计', '果园', '茶园', '其他园地',
              '小计', '有林', '灌木林地', '其他林地',
              '小计', '天然', '人工', '其他',
              '小计', '城市', '建制', '村庄', '采矿用地', '风景',
              '小计', '铁路用地', '公路用地', '农村道路', '机场用地', '港口码头用地', '管道运输用地',
              '小计', '河流水面', '湖泊水面', '水库水面', '坑塘水面', '沿海滩涂', '内陆滩涂', '沟渠', '水工建筑用地', '冰川',
              '小计', '空闲', '设施', '田坎', '盐碱', '沼泽', '沙地', '裸地'
              ]

    row4_mc = [[1, 0, 1, 1, '面积:公顷'],
               [2, 0, 4, 0, '序号'],
               [2, 1, 4, 1, '政码'],
               [2, 2, 4, 2, '行政区名'],
               [2, 3, 4, 3, '总计'],
               [2, 4, 2, 7, '耕地'],
               [2, 8, 2, 11, '园地'],
               [2, 12, 2, 15, '林地'],
               [2, 16, 2, 19, '草地'],
               [2, 20, 2, 25, '城镇村及工矿用地'],
               [2, 26, 2, 32, '交通运输用'],
               [2, 33, 2, 42, '水域及水利设施用'],
               [2, 43, 2, 50, '其他土地']
               ]

    worksheet.write_row('E4', row_mc)
    for cell in row4_mc:
        # worksheet.write_row(cell[0],cell[1])
        worksheet.merge_range(cell[0], cell[1], cell[2], cell[3], cell[4], classify)
    worksheet.merge_range(0,0,0,50,"地块按土地利用现状分类面积统计汇总表",title)
    # worksheet.write_row('E3', ['耕地'])
    # worksheet.write_row('I3', ['园地'])
    # worksheet.write_row('M3', ['林地'])
    # worksheet.write_row('Q3', ['草地'])
    # worksheet.write_row('U3', ['城镇村及工矿用地'])
    # worksheet.write_row('AA3', ['交通运输用地'])
    # worksheet.write_row('AH3', ['水域及水利设施用地'])
    # worksheet.write_row('AR3', ['其他用地'])
	


def write_row_bm(row_bm):
    worksheet.write(row_bm, 4, '01')
    worksheet.write(row_bm, 5, '011')
    worksheet.write(row_bm, 6, '012')
    worksheet.write(row_bm, 7, '013')
    worksheet.write(row_bm, 8, '02')
    worksheet.write(row_bm, 9, '021')
    worksheet.write(row_bm, 10, '022')
    worksheet.write(row_bm, 11, '023')
    worksheet.write(row_bm, 12, '03')
    worksheet.write(row_bm, 13, '031')
    worksheet.write(row_bm, 14, '032')
    worksheet.write(row_bm, 15, '033')


def get_zldwdm(rows):
    zldwdm_list = []
    for record in rows:
        if record[0] not in zldwdm_list:
            zldwdm_list.append(record[0])
    return zldwdm_list


def write_into_xlsx():
    row = 4
    col = 0
    zldwdm_list = []
    for record in records:
        # print record
        zldwdm = record[0]
        dlbm = record[1]
        dlmj = record[2]
        if zldwdm not in zldwdm_list:
            zldwdm_list.append(zldwdm)
            row = row + 1
        worksheet.write(row, col, row - 4)
        worksheet.write(row, col + 1, zldwdm)
        if dlbm.strip() == '01':
            worksheet.write(row, col + 4, dlmj)
        if dlbm.strip() == '011':
            worksheet.write(row, col + 5, dlmj)
        if dlbm.strip() == '012':
            worksheet.write(row, col + 6, dlmj)
        if dlbm.strip() == '013':
            worksheet.write(row, col + 7, dlmj)
        if zldwdm not in zldwdm_list:
            row = row  + 1
            zldwdm_list.append(zldwdm)


def main():
    write_dlmc_row()
    write_row_bm(4)
    write_into_xlsx()
    workbook.close()
main()
