# coding: utf-8
import os
import sys
import time
import pypyodbc
import xlsxwriter
import arcgisscripting as arc
reload(sys)
sys.setdefaultencoding('utf-8')


class WriteToExcel:

    def __init__(self, path = None):
        file = os.path.basename(path).split('.')[0]
        filename = os.path.dirname(path) +'/' + file + '.xlsx'
        dic1 = {'bold': 0, 'text_v_align': 2, 'text_h_align': 2}
        dic2 = {'bold': 1, 'text_h_align': 2, 'text_v_align': 2, 'font_name': '宋体', 'font_size': 16}
        self.row = 6
        self.title_format = self.workbook.add_format(dic2)
        self.classify_format = self.workbook.add_format(dic1)
        self.cell_format = self.workbook.add_format({'text_v_align': 2,'text_h_align': 2})
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet_base = self.workbook.add_worksheet('基期地类')
        self.worksheet_ghdl = self.workbook.add_worksheet('规划地类')
        self.worksheet_gzq = self.workbook.add_worksheet('占用管制区')
        self.worksheet_ghyt = self.workbook.add_worksheet('占用规划用途')
        self.gzq_dm = {'zmj': 1, '01': 2, '01g': 3, '01j': 4,
                 '02': 5, '02g': 6, '02j': 7,
                 '03': 8, '03g': 9, '03j': 10,
                 '04': 11, '04g': 12, '04j': 13
                 }
        self.jq_dm = {'zmj': 3, '1': 4, '11': 5, '12': 6, '13': 7, '14': 8, '15': 9,
                    '151': 10, '152': 11, '153': 12, '154': 13, '155': 14,
                    '2': 15, '21': 16, '211': 17, '212': 18, '213': 19, '214': 20,
                    '22': 21, '221': 22, '222': 23, '223': 24, '224': 25, '225': 26, '226': 27, '227': 28,
                    '23': 29, '231': 30, '232': 31, '233': 32,
                    '3': 33, '31': 34, '311': 35, '312': 36, '313': 37, '32': 38
                    }
        self.ghyt_dm = {'zmj':1,'1':2,'G111': 3, 'G112': 4, 'N111': 5, 'N112': 6, 'X12': 7, 'G12': 8, 'X13': 9,
                    'G13': 10, 'X14': 11, 'G14': 12, '15': 13, '2': 14,
                    '21': 15, 'X211': 16, 'X212': 17, 'X213': 18, 'X214': 19, 'G211': 20,
                    'G212': 21, 'G213': 22, 'G214': 23, '22': 24, 'X221': 25, 'X222': 26, 'X223': 27, 'X224': 28,
                    'X225': 29, 'X226': 30, 'X227': 31, 'G221': 32,
                    'G222': 33, 'G223': 34, 'G224': 35, 'G225': 36, 'G226': 37, 'G227': 38,
                    '23': 39, 'X231': 40, 'X232': 41, 'X233': 42, 'G231': 43, 'G232': 44, 'G233': 45,
                    '3': 46, '31': 47, 'X311': 48, 'X312': 49, 'X313': 50, 'G311': 51, 'G312': 52, 'G313': 53, 'X32': 54
                   }

				   
    def write_dlmc_row(self, worksheet):

        jqdlmc = ['', '', '', '',
                  '小计', '设施农用地', '农村道路', '坑塘水面', '农田水利用地', '田坎', '',
                  '小计', '城镇用地', '农村居民点用地', '采矿用地', '其他独立建设用地',
                  '小计', '铁路用地', '公路用地', '民用机场用地', '港口码头用地', '管道运输用地', '水库水面', '水工建筑用地',
                  '小计', '风景名胜建设用地', '特殊用地', '盐田', '',
                  '小计', '河流水面', '湖泊水面', '滩涂沼泽', '(32)',
                  ]
        jqdldm = ['(11)', '(12)', '(13)', '(14)', '',
                  '(151)', '(152)', '(153)', '(154)', '(155)', '', '',
                  '(211)', '(212)', '(213)', '(214)', '',
                  '(221)', '(222)', '(223)', '(224)', '225', '226', '(227)', '',
                  '(231)', '(232)', '(233)', '', '',
                  '(311)', '(312)', '(313)', '(32)']
        jqmc = [[2, 4, 2, 14, '农用地'],
                   [2, 15, 2, 32, '建设用地'],
                   [2, 33, 2, 38, '未利用地'],
                   [2, 0, 5, 0, '地 块'],
                   [2, 1, 5, 1, '行政区名称'],
                   [2, 2, 5, 2, '行政区代码'],
                   [2, 3, 5, 3, '总 计'],
                   [3, 4, 5, 4, '合 计'],
                   [3, 5, 4, 5, '耕 地(11)'],
                   [3, 6, 4, 6, '园 地(12)'],
                   [3, 7, 4, 7, '林 地(13)'],
                   [3, 8, 4, 8, '牧草地(14)'],
                   [3, 9, 3, 14, '其他农用地(15)'],
                   [3, 15, 5, 15, '合 计'],
                   [3, 16, 3, 20, '城乡建设用地(21)'],
                   [3, 21, 3, 28, '交通水利用地(22)'],
                   [3, 29, 3, 32, '其他建设用地(23)'],
                   [3, 33, 5, 33, '合 计'],
                   [3, 34, 3, 37, '水域(31)'],
                   [3, 38, 4, 38, '自然保留地(32)']
                   ]
        ytdldm = ['G111', 'G112', 'N111', 'N112', 'X12', 'G12',
                'X13',  'G13', 'X14', 'G14', '15', '', '',
                'X211', 'X212', 'X213', 'X214', 'G211', 'G212', 'G213', 'G214', '',
                'X221', 'X222', 'X223', 'X224', 'X225', 'X226', 'X227', 'G221',
                'G222', 'G223', 'G224', 'G225', 'G226', 'G227', '',
                'X231', 'X232', 'X233', 'G231', 'G232', 'G233', '', '',
                'X311', 'X312', 'X313', 'G311', 'G312', 'G313', 'X32'
                ]
        ytdlmc = ['示范区基本农田', '一般基本农田', '一般农田', '新增一般农田', '园地', '新增园地',
                '林地', '新增林地', '牧草地', '新增牧草地', '其他农用地', '', '小计',
                '城镇用地', '农村居民点用地', '采矿用地', '其他独立建设用地', '新增城镇用地', '新增农村居民点用地',
                '新增采矿用地', '新增其他独立建设用地', '小计',
                '铁路用地', '公路用地', '机场用地', '港口码头用地', '管道运输用地', '水库水面', '水共建筑用地', '新增铁路用地',
                '新增公路用地', '新增民用机场用地', '新增港口码头用地', '新增管道运输用地', '新增水库水面', '新增水共建筑用地',
                '小计', '风景名胜设施用地', '特殊用地', '盐田', '新增风景名胜设施用地', '新增特殊用地', '新增盐田', '',
                '小计', '河流水面', '湖泊水面', '滩涂', '新增河流水面', '新增滩涂', '新增湖泊水面', ''
                ]
        ytmc = [[2, 0, 5, 0 , '地 块'],
                [2, 1, 5, 1, '总 计'],
                [2, 2, 2, 13, '农用地(1)'],
                [2, 14, 2, 45, '建设用地(2)'],
                [2, 46, 2, 54, '其他土地(3)'],
                [3, 2, 5, 2, '合 计'],
                [3, 3, 3, 4, '基本农田(G11)'],
                [3, 5, 3, 6, '一般农田(N11)'],
                [3, 7, 3, 8, '园地(12)'],
                [3, 9, 3, 10, '林地(13)'],
                [3, 11, 3, 12, '牧草地(14)'],
                [3, 13, 4, 13, '其他农用地'],
                [3, 14, 5, 14, '合 计'],
                [3, 15, 3, 23, '城乡建设用地(21)'],
                [3, 24, 3, 38, '交通水利用地(22)'],
                [3, 39, 3, 45, '其他建设用地(23)'],
                [3, 46, 5, 46, '合 计'],
                [3, 47, 3, 53, '水域(31)'],
                [3, 54, 4, 54, '自然保留地']
                ]
        gzdlmc = ['小 计', '其中耕地', '基本农田',
                '小 计', '其中耕地', '基本农田',
                '小 计', '其中耕地', '基本农田',
                '小 计', '其中耕地', '基本农田'
                ]
        gzmc = [[2, 0, 3, 0, '地 块'],
                [2, 1, 3, 1, '总面积'],
                [2, 2, 2, 4, '允许建设区'],
                [2, 5, 2, 7, '有条件建设区'],
                [2, 8, 2, 10, '限制建设区'],
                [2, 11, 2, 13, '禁止建设区']
                ]

        if worksheet is self.worksheet_ghdl:
            title = "规划地类面积汇总表"
            worksheet.write_row('F6', jqdldm, self.classify_format)
            worksheet.write_row('F5', jqdlmc, self.classify_format)
            for cell in jqmc:
                worksheet.merge_range(cell[0], cell[1], cell[2], cell[3], cell[4], self.classify_format)
            worksheet.merge_range(1, 2, 1, 38, '')
            worksheet.merge_range(0, 0, 0, 38, title, self.title_format)
            worksheet.set_row(4, 25)
            worksheet.set_row(5, 20)
        elif worksheet is self.worksheet_base:
            title = "基期地类面积统计汇总表"
            worksheet.write_row('F6', jqdldm, self.classify_format)
            worksheet.write_row('F5', jqdlmc, self.classify_format)
            for cell in jqmc:
                worksheet.merge_range(cell[0], cell[1], cell[2], cell[3], cell[4], self.classify_format)
            worksheet.merge_range(1, 2, 1, 38, '')
            worksheet.merge_range(0, 0, 0, 38, title, self.title_format)
            worksheet.set_row(4, 25)
            worksheet.set_row(5, 20)
        elif worksheet is self.worksheet_ghyt:
            title = "项目占用规划用途面积"
            worksheet.write_row('D6', ytdldm, self.classify_format)
            worksheet.write_row('D5', ytdlmc, self.classify_format)
            for cell in ytmc:
                worksheet.merge_range(cell[0], cell[1], cell[2], cell[3], cell[4], self.classify_format)
            worksheet.merge_range(1, 2, 1, 54, '')
            worksheet.merge_range(0, 0, 0, 54, title, self.title_format)
        else:
            title = "项目占用建设用地管制区"
            worksheet.write_row('C4', gzdlmc, self.classify_format)
            for cell in gzmc:
                worksheet.merge_range(cell[0], cell[1], cell[2], cell[3], cell[4], self.classify_format)
            worksheet.merge_range(1, 2, 1, 13, '')
            worksheet.merge_range(0, 0, 0, 13, title, self.title_format)
        if is_hectare == 'true':
            worksheet.merge_range(1, 0, 1, 1, '面积单位: 公顷(0.0000)', self.classify_format)
        else:
            worksheet.merge_range(1, 0, 1, 1, '面积单位: 平方米(0.00)', self.classify_format)
        # 设置表头格式
        worksheet.set_row(0, 30)
        worksheet.set_row(1, 25)
        worksheet.set_row(2, 30)
        worksheet.set_row(3, 30)


    def write_to_xlsx(self, worksheet, records, dldm_v, r, c):
        """
        jqtb: r = 6, c = 40
        ghdl: r = 6, c = 40
        ghyt: r = 6, c = 55
        gzq: r = 5, c = 13
        """
        row = r
        col = c
        hide_row = 2
        col_list = []
        zldwdm_list = []
        for record in records:
            zldwdm = record[0]
            dlbm = record[1]
            dlmj = record[2]

            if zldwdm not in zldwdm_list:
                zldwdm_list.append(zldwdm)
                row += 1
            if len(record) == 4:
                zldwmc = record[3]
                worksheet.write(row, 1, zldwmc)
                worksheet.write(row, 2, zldwdm)
            else:
                worksheet.write(row, 0, zldwdm)
            for dic in dldm_v:
                if dic == dlbm:
                    col_v = dldm_v[dic]
                    worksheet.write(row, col_v, dlmj)
                    col_list.append(col_v)
                continue
            worksheet.set_row(row, 25, self.cell_format)
        # 隐藏没有数据的列
        if len(records[0]) == 4:
            hide_row = 5
        for i in range(hide_row, col, 1):
            if i not in col_list:
                worksheet.set_column(i, i, 0)


    def close(self):
        self.workbook.close()


def update_features(in_fc):
    """
    处理各个图层的面积、田坎系数等字段值
    """
    rows = GP.UpdateCursor(in_fc)
    row = rows.Next()
    while row:
        if in_fc == "PDLTB":
            if row.TKXS > 1:
                row.TKXS = (row.TKXS/100)
            if is_hectare == "true":
                row.TBMJ = round((row.shape_Area/10000), 4)
                row.KKSM = round(row.TBMJ*row.TKXS, 4)
            else:
                row.TBMJ = round(row.shape_Area, 2)
                row.KKSM = round(row.TBMJ*row.TKXS, 2)
        elif in_fc == "PXZDW":
            row.XWSC = round(row.SHAPE_LENGTH, 1)
            if is_hectare == "true":
                row.XZDWMJ = round(row.XWSC*row.XWKD/10000, 4)
            else:
                row.XZDWMJ = round(row.XWSC*row.XWKD, 2)
        elif in_fc == "PLXDW":
            if is_hectare == "true":
                row.LXDWMJ = round(row.LXDWMJ/10000, 4)
        elif in_fc == "PGHYT":
            if is_hectare == "true":
                row.MJ = round((row.shape_Area/10000), 4)
            else:
                row.MJ = round(row.shape_Area, 2)
        elif in_fc == "PGZQ":
            if is_hectare == "true":
                row.GZQMJ = round((row.shape_Area/10000), 4)
            else:
                row.GZQMJ = round(row.shape_Area, 2)
        elif in_fc == "PGHDL":
            if is_hectare == "true":
                row.GHDLMJ = round((row.shape_Area/10000), 4)
            else:
                row.GHDLMJ = round(row.shape_Area, 2)
        rows.updaterow(row)
        row = rows.Next()


def overlay():
    gh_features = ["PXZDW", "PDLTB", "PLXDW", "PGHYT", "PGZQ", "PGHDL", "DK"]
    for fc in gh_features:
        if GP.exists(fc):
            GP.delete_management(fc)
    GP.CopyFeatures_management(input_dk, "DK")
    in_tb = 'JQDLTB' + ';' + "DK"
    in_xw = 'JQXZDW' + ';' + "DK"
    in_lw = 'JQLXDW' + ';' + "DK"
    in_ghyt = 'GHYT' + ';' + "DK"
    in_ghdl = 'TDGHDL' + ';' + "DK"
    in_gzq = 'JSYDGZQ' + ';' + "DK"
    GP.AddMessage("Intersecting ...")
    GP.Intersect_analysis(in_tb, "PDLTB", "ALL", "", "")
    GP.Intersect_analysis(in_ghyt, "PGHYT", "ALL", "", "")
    GP.Intersect_analysis(in_ghdl, "PGHDL", "ALL", "", "")
    GP.Intersect_analysis(in_gzq, "TGZQ", "ALL", "")
    GP.Intersect_analysis(in_xw, "TXZDW", "ALL", "", "")
    GP.Intersect_analysis(in_lw, "TLXDW", "ALL", "", "")
    GP.AddMessage("Intersect successfully!" + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    update_features("PDLTB")
    update_features("PGHYT")
    update_features("PGHDL")
    GP.identity("TXZDW", "PDLTB", "PXZDW", "ALL", "", "KEEP_RELATIONSHIPS")
    GP.identity("TLXDW", "PDLTB", "PLXDW", "ALL", "", "")
    GP.identity("TGZQ", "PGHYT", "PGZQ", "ALL", "", "")
    GP.AddField("PGHYT", "DK_FID", "TEXT", "", "", "", "DK", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddMessage("Identity analysis successfully!" + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    update_features("PGZQ")
    update_features("PXZDW")
    update_features("PLXDW")
    GP.delete_management("TXZDW")
    GP.delete_management("TLXDW")
    GP.delete_management("TGZQ")


def create_table():
    t_table = ["T_HZMJ", "T_XZDW", "T_LXDW", "T_DLTB", "T_GHYT", "T_GZQ", "T_GHDL"]
    for table in t_table:
        if GP.exists(table):
            GP.delete_management(table)
    GP.CreateTable_management(location, "T_HZMJ")
    # Add Fields To Table "T_HZMJ"
    GP.AddField("T_HZMJ", "ZLDWDM", "TEXT", "", "", "", "ZLDWDM", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField("T_HZMJ", "ZLDWMC", "TEXT", "", "", "", "ZLDWMC", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField("T_HZMJ", "QSDWDM", "TEXT", "", "", "", "QSDWDM", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField("T_HZMJ", "QSDWMC", "TEXT", "", "", "", "QSDWMC", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField("T_HZMJ", "QSXZ", "TEXT", "", "", "", "QSXZ", "NULLABLE", "REQUIRED", "#")
    GP.AddField("T_HZMJ", "DLBM", "TEXT", "", "", "8", "DLBM", "NULLABLE", "REQUIRED", "#")
    # GP.AddField("T_HZMJ","DLMC","TEXT","","","","DLMC","NULLABLE","REQUIRED","#")
    # GP.AddField("T_HZMJ","TBBH","TEXT","","","","TBBH","NULLABLE","NON_REQUIRED","#")
    # GP.AddField("T_HZMJ","TBMJ","DOUBLE","","","","TBMJ","NULLABLE","NON_REQUIRED","#")
    # GP.AddField("T_HZMJ","LXDWMJ","DOUBLE","","","","LXDWMJ","NULLABLE","NON_REQUIRED","#")
    # GP.AddField("T_HZMJ","XZDWMJ","DOUBLE","","","","XZDWMJ","NULLABLE","NON_REQUIRED","#")
    # GP.AddField("T_HZMJ","TKMJ","DOUBLE","","","#", "TKMJ", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField("T_HZMJ", "DLMJ", "DOUBLE", "", "", "#", "DLMJ", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField("T_HZMJ", "TABLENAME", "TEXT", "", "", "", "TABLENAME", "NULLABLE", "REQUIRED", "#")
    GP.AddField("T_HZMJ", "flag", "TEXT", "", "", "#", "flag", "NULLABLE", "NON_REQUIRED", "#")


def data_statistic():
    conn = pypyodbc.connect('Driver={Microsoft Access Driver (*.mdb)};DBQ=' + location)
    cur = conn.cursor()
    cur.execute('''UPDATE PDLTB SET KSXM=0,KLWM=0''')
    cur.commit()
    cur.execute('''update pdltb a,pxzdw b set a.KSXM=a.KSXM +b.xzdwmj*b.kcxs
    where a.objectid = b.left_pdltb''')
    cur.commit()
    cur.execute('''update pdltb a,pxzdw b set a.KSXM=a.KSXM +b.xzdwmj*(1-b.kcxs)
    where a.objectid = b.right_pdltb''')
    cur.commit()
    cur.execute('''update pdltb a,plxdw b set a.KLWM =a.KLWM +b.LXDWMJ where a.objectid = b.fid_pdltb ''')
    cur.commit()
    cur.execute('''update pdltb set TBDLMJ = TBMJ-KSXM-KLWM-KKSM ''')
    cur.commit()
    GP.AddMessage("Update KSXM KLWM KLWM finished " +
                  time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    # INSERT INTO TABLE T_LXDW
    cur.execute('''select DLDM,SZTBBH as tbbh,qsxz,left(XZQDM,12) as TZLDWDM,XZQMC AS ZLDWMC,
    left(SQBM,12) as TQSDWDM,SQMC as QSDWMC,lxdwmj,'lxdw' as tablename into t_lxdw from plxdw''')
    # INSERT INTO  TABLE T_XZDW
    cur.execute('''select a.DLDM,a.kctbbh1 as tbbh,a.qsxz,left(a.KCXZQDM1,12) as ZLDWDM,'' as zldwmc,
    left(a.SQBM1,12) as qsdwdm,a.SQMC1 as qsdwmc,sum(a.xzdwmj*a.KCXS) as xzdwmj,
    'xzdw' as tablename into t_xzdw from pxzdw a,pdltb b where a.left_pdltb = b.objectid group by
    a.DLDM,a.kctbbh1,a.qsxz,left(a.KCXZQDM1,12),'',left(a.SQBM1,12),a.SQMC1''')
    # ADD TO TABLE T_XZDW
    cur.execute('''insert into t_xzdw(DLDM,tbbh,qsxz,zldwdm,zldwmc,qsdwdm,qsdwmc,xzdwmj,tablename)
    select a.DLDM,a.kctbbh2,a.qsxz,left(a.KCXZQDM2,12),'',left(a.SQBM2,12),a.SQMC2,
    sum(a.xzdwmj*(1-a.KCXS)),'xzdw' from pxzdw a,pdltb b where  a.right_pdltb = b.objectid group by
    a.DLDM, a.kctbbh2,a.qsxz,left(a.KCXZQDM2,12),'',left(a.SQBM2,12),a.SQMC2''')
    cur.commit()
    #####
    # INSERT INTO T_DLTB
    cur.execute('''select DLDM,tbbh,qsxz,left(XZQDM,12) as TZLDWDM,XZQMC AS ZLDWMC,left(SQBM,12) as TQSDWDM,SQMC AS QSDWMC,
    sum(tbmj) as ttbmj,sum(tbdlmj) as ttbdlmj,sum(KSXM) as xwmj,sum(KLWM) as lwmj,sum(KKSM) as ttkmj,
    'dltb' as tablename into t_dltb from pdltb group by DLDM,tbbh,qsxz,left(XZQDM,12),XZQMC,left(SQBM,12),SQMC''')
    cur.commit()
    GP.AddMessage("Finish insert into t_dltb")
    # INSERT INTO T_GHDL
    cur.execute('''select left(xzqdm,12) as zldwdm,xzqmc as zldwmc,ghdldm,sum(ghdlmj) as dlmj,'3' as flag into t_ghdl
    from PGHDL group by left(xzqdm,12),xzqmc,ghdldm''')
    cur.execute('''insert into t_ghdl(zldwdm,zldwmc,ghdldm,dlmj,flag) select left(xzqdm,12),xzqmc,left(ghdldm,2),
    sum(ghdlmj),'2' from PGHDL group by left(xzqdm,12),xzqmc,left(ghdldm,2)''')
    cur.execute('''insert into t_ghdl(zldwdm,zldwmc,ghdldm,dlmj,flag) select left(xzqdm,12),xzqmc,left(ghdldm,1),
    sum(ghdlmj),'1' from PGHDL group by left(xzqdm,12),xzqmc,left(ghdldm,1)''')
    cur.execute('''insert into t_ghdl(zldwdm,zldwmc,ghdldm,dlmj,flag) select zldwdm,zldwmc,'zmj',sum(dlmj),'0' from
    t_ghdl where flag='1' group by zldwdm,zldwmc ''')
    cur.commit()
    # INSERT INTO T_GHYT
    cur.execute('''select fid_dk,ghytdm,sum(mj) as dlmj,'3' as flag into t_ghyt from PGHYT group by fid_dk,ghytdm ''')
    cur.execute('''insert into t_ghyt(fid_dk,ghytdm,dlmj,flag)select fid_dk,left(ghytdm,3),sum(mj),'2' from
    PGHYT group by fid_dk,left(ghytdm,3)''')
    cur.execute('''insert into t_ghyt(fid_dk,ghytdm,dlmj,flag) select fid_dk,right(left(ghytdm,3),2),sum(mj),'2' from
    PGHYT group by fid_dk,right(left(ghytdm,3),2)''')
    cur.execute('''insert into t_ghyt(fid_dk,ghytdm,dlmj,flag)select fid_dk,right(left(ghytdm,2),1),sum(mj),'1' from
    PGHYT group by fid_dk,right(left(ghytdm,2),1)''')
    cur.execute('''insert into t_ghyt(fid_dk,ghytdm,dlmj,flag)select fid_dk,'zmj',sum(dlmj),'0' from t_ghyt where
    flag = '1' group by fid_dk''')
    cur.commit()
    # INSERT INTO T_GZQ
    cur.execute('''select fid_dk_1 as fid_dk, left(gzqlxdm, 2) as gzqdm, sum(gzqmj) as mj, left(gzqlxdm, 2) as flag into
    t_gzq from PGZQ group by fid_dk_1, left(gzqlxdm, 2) ''')
    cur.execute('''insert into t_gzq(fid_dk, gzqdm, mj, flag) select fid_dk_1, left(gzqlxdm, 2), sum(gzqmj),
    left(gzqlxdm, 2)+'g' from PGZQ where ghytdm in('G111', 'G112', 'N111', 'N112') group by fid_dk_1, left(gzqlxdm, 2)''')
    cur.execute('''insert into t_gzq(fid_dk, gzqdm, mj, flag) select fid_dk_1, left(gzqlxdm, 2), sum(gzqmj),
    left(gzqlxdm, 2)+'j' from PGZQ where ghytdm in('G111', 'G112') group by fid_dk_1, left(gzqlxdm, 2) ''')
    cur.execute('''insert into t_gzq(fid_dk,mj,flag) select fid_dk_1, sum(gzqmj), 'zmj' from PGZQ group by fid_dk_1''')
    cur.commit()
    #####
    #INSERT INTO T_HZMJ
    #####
    cur.execute('''insert into T_HZMJ(zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,dlmj,tablename) select tzldwdm,zldwmc,tqsdwdm,
    qsdwmc, DLDM,sum(lxdwmj),tablename from t_lxdw group by tzldwdm,zldwmc,tqsdwdm,qsdwmc,DLDM,tablename''')
    cur.commit()
    cur.execute('''insert into T_HZMJ(zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,dlmj,tablename) select zldwdm,zldwmc,qsdwdm,
    qsdwmc, DLDM,sum(xzdwmj),tablename from t_xzdw where xzdwmj>0 group by zldwdm,zldwmc,qsdwdm,qsdwmc,DLDM,tablename''')
    cur.commit()
    cur.execute('''insert into T_HZMJ(zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,dlmj,tablename) select tzldwdm,zldwmc,tqsdwdm,
    qsdwmc, DLDM,sum(ttbdlmj),tablename from t_dltb group by tzldwdm,zldwmc,tqsdwdm,qsdwmc,DLDM,tablename''')
    cur.commit()
    cur.execute('''insert into T_HZMJ(zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,dlmj,tablename) select tzldwdm,zldwmc,tqsdwdm,
    qsdwmc,'155', sum(ttkmj), 'tk' from t_dltb where ttkmj>0 group by tzldwdm,zldwmc,tqsdwdm,qsdwmc''')
    #create temp table tt to update zldwmc
    cur.execute('''select distinct(zldwmc), zldwdm into tt from "T_HZMJ" where zldwmc is not null and zldwmc <>'' ''')
    cur.execute('''update "T_HZMJ" a, tt b set a.zldwmc = b.zldwmc where a.zldwdm = b.zldwdm ''')
    cur.commit()
    #update zldwmc successfully
    #insert into "T_HZMJ" continue
    cur.execute('''insert into T_HZMJ(zldwdm,zldwmc,dlbm,dlmj,flag) select zldwdm,zldwmc,dlbm,sum(dlmj),'3' from
    "T_HZMJ" where len(dlbm)=3 group by zldwdm,zldwmc,dlbm''')
    cur.execute('''insert into T_HZMJ(zldwdm,zldwmc,dlbm,dlmj,flag) select zldwdm,zldwmc,left(dlbm,2),sum(dlmj),'2'
    from "T_HZMJ" where tablename in('lxdw','xzdw','dltb','tk') group by zldwdm,zldwmc,left(dlbm,2) ''')
    cur.execute('''insert into T_HZMJ(zldwdm,zldwmc,dlbm,dlmj,flag) select zldwdm,zldwmc,left(dlbm,1),sum(dlmj),'1'
    from "T_HZMJ" where flag = '2' group by zldwdm,zldwmc,left(dlbm,1) ''')
    cur.execute('''insert into T_HZMJ(zldwdm,zldwmc,dlbm,dlmj,flag) select zldwdm,zldwmc,'zmj',sum(dlmj),'0' from
    "T_HZMJ" where flag = '1' group by zldwdm,zldwmc ''')
    cur.execute('''drop table tt''')
    cur.execute('''insert into "T_HZMJ"(dlbm,dlmj) select dlbm,sum(dlmj) from T_HZMJ where flag = '2'
    group by dlbm ''')
    cur.execute('''insert into "T_HZMJ"(dlbm,dlmj) select dlbm,sum(dlmj) from T_HZMJ where flag = '1'
    group by dlbm''')
    cur.execute('''insert into "T_HZMJ"(dlbm,dlmj) select 'zmj',sum(dlmj) from T_HZMJ  where flag ='2' ''')
    cur.commit()
    # jqdl
    toExecl = WriteToExcel(location)
    data = cur.execute('''select zldwdm,dlbm,dlmj,zldwmc from T_HZMJ where flag in('0','1','2','3') order by zldwdm ''')
    records_jqdl = data.fetchall()
    toExecl.write_dlmc_row(toExecl.worksheet_base)
    toExecl.write_to_xlsx(toExecl.worksheet_base, records_jqdl, toExecl.jq_dm, 6, 40)
    # ghdl
    data_ghdl = cur.execute('''select zldwdm,ghdldm,dlmj,zldwmc from t_ghdl where flag in('0','1','2','3') order by zldwdm ''')
    records_ghdl = data_ghdl.fetchall()
    toExecl.write_dlmc_row(toExecl.worksheet_ghdl)
    toExecl.write_to_xlsx(toExecl.worksheet_ghdl, records_ghdl, toExecl.jq_dm, 6, 40)
    # ghyt
    data_ghyt = cur.execute('''select fid_dk,ghytdm,dlmj from t_ghyt where flag in('0','1','2','3') order by fid_dk''')
    records_ghyt = data_ghyt.fetchall()
    toExecl.write_dlmc_row(toExecl.worksheet_ghyt)
    toExecl.write_to_xlsx(toExecl.worksheet_ghyt, records_ghyt, toExecl.ghyt_dm, 6, 55)
    # gzq
    data_gzq = cur.execute('''select fid_dk, flag, mj from t_gzq order by fid_dk''')
    records_gzq = data_gzq.fetchall()
    toExecl.write_dlmc_row(toExecl.worksheet_gzq)
    toExecl.write_to_xlsx(toExecl.worksheet_gzq, records_gzq, toExecl.gzq_dm, 4, 14)

    toExecl.close()
    cur.close()
    conn.close()


def main():
    overlay()
    create_table()
    data_statistic()


if __name__ == '__main__':
    GP = arc.create(9.3)
    location = GP.GetParameterAsText(0)
    input_dk = GP.GetParameterAsText(1)
    input_ptb = GP.GetParameterAsText(2)
    is_hectare = GP.GetParameterAsText(3)
    GP.Workspace = location
    main()
