# coding:utf-8
import sys
import os
import time
import arcgisscripting as arc
import pypyodbc
import xlsxwriter
reload(sys)
sys.setdefaultencoding('utf-8')


class ManipulateExcel:

    def __init__(self, filename = None, path = None):
        self.row = 4
        self.col = 0
        self.col_list = []
        self.zldwdmlist = []
        self.workbook = xlsxwriter.Workbook(os.path.dirname(path) + '/' + filename + '.xlsx')
        self.worksheet = self.workbook.add_worksheet(filename)
        self.dic1 = {'bold': 0, 'text_v_align': 2, 'text_h_align': 2}
        self.dic2 = {'bold': 1, 'text_h_align': 2, 'text_v_align': 2,
                'font_name': '宋体', 'font_size': 16
                }
        self.row_bm = ['01', '011', '012', '013',
                  '02', '021', '022', '023',
                  '03', '031', '032', '033',
                  '04', '041', '042', '043',
                  '20', '201', '202', '203', '204', '205',
                  '10', '101', '102', '104', '105', '106', '107',
                  '11', '111', '112', '113', '114', '115', '116', '117', '118', '119',
                  '12', '121', '122', '123', '124', '125', '126', '127'
                  ]
        self.row_mc = ['小计', '水田', '水浇地', '旱地',
                  '小计', '果园', '茶园', '其他园地',
                  '小计', '有林地', '灌木林地', '其他林地',
                  '小计', '天然牧草地', '人工牧草地', '其他草地',
                  '小计', '城市', '建制镇', '村庄', '采矿用地', '风景名胜及特殊用地',
                  '小计', '铁路用地', '公路用地', '农村道路', '机场用地', '港口码头用地', '管道运输用地',
                  '小计', '河流水面', '湖泊水面', '水库水面', '坑塘水面', '沿海滩涂', '内陆滩涂', '沟渠', '水工建筑用地', '冰川及永久积雪',
                  '小计', '空闲地', '设施农用地', '田坎', '盐碱地', '沼泽地', '沙地', '裸地'
                  ]
        self.row4_mc = [[2, 0, 4, 0, '序 号'],
                   [2, 1, 4, 1, '行政区代码'],
                   [2, 2, 4, 2, '行政区名称'],
                   [2, 3, 4, 3, '总 计'],
                   [2, 4, 2, 7, '耕 地'],
                   [2, 8, 2, 11, '园 地'],
                   [2, 12, 2, 15, '林 地'],
                   [2, 16, 2, 19, '草 地'],
                   [2, 20, 2, 25, '城镇村及工矿用地'],
                   [2, 26, 2, 32, '交通运输用地'],
                   [2, 33, 2, 42, '水域及水利设施用地'],
                   [2, 43, 2, 50, '其他土地']
                   ]
        self.dic_dlbm = {'zmj': 3, '01': 4, '011': 5, '012': 6, '013': 7, '02': 8, '021': 9, '022': 10, '023': 11,
                    '03': 12, '031': 13, '032': 14, '033': 15, '04': 16, '041': 17, '042': 18, '043': 19,
                    '20': 20, '201': 21, '202': 22, '203': 23, '204': 24, '205': 25,
                    '10': 26, '101': 27, '102': 28, '104': 29, '105': 30, '106': 31, '107': 32,
                    '11': 33, '111': 34, '112': 35, '113': 36, '114': 37, '115': 38, '116': 39, '117': 40, '118': 41, '119': 42,
                    '12': 43, '121': 44, '122': 45, '123': 46, '124': 47, '125': 48, '126': 49, '127': 50
                    }


    def write_dlmc_row(self):
        classify_format = self.workbook.add_format(self.dic1)
        title_format = self.workbook.add_format(self.dic2)
        self.worksheet.write_row('E4', self.row_mc, classify_format)
        self.worksheet.write_row('E5', self.row_bm, classify_format)
        for cell in self.row4_mc:
            self.worksheet.merge_range(cell[0], cell[1], cell[2], cell[3], cell[4], classify_format)
        self.worksheet.merge_range(0, 0, 0, 50, "地块按土地利用现状分类面积统计汇总表", title_format)
        if is_hectare == 'true':
            self.worksheet.merge_range(1, 0, 1, 1, '面积: 公顷', classify_format)
        else:
            self.worksheet.merge_range(1, 0, 1, 1, '面积: 平方米', classify_format)
        # 设置表头格式
        self.worksheet.merge_range(1, 2, 1, 50, '')
        self.worksheet.set_row(0, 30)
        self.worksheet.set_row(1, 30)
        self.worksheet.set_row(2, 30)
        self.worksheet.set_row(3, 30)
        self.worksheet.set_row(4, 20)


    def write_to_xlsx(self, records):
        """
        生成Excel表
        """
        self.write_dlmc_row()
        cell_format = self.workbook.add_format({'text_v_align': 2,'text_h_align': 2})
        for record in records:
            zldwdm = record[0]
            zldwmc = record[1]
            dlbm = record[2]
            dlmj = record[3]
            if zldwdm not in self.zldwdmlist:
                self.zldwdmlist.append(zldwdm)
                self.row = self.row + 1
            self.worksheet.write(self.row, self.col, self.row - 4)
            self.worksheet.write(self.row, 1, zldwdm)
            self.worksheet.write(self.row, 2, zldwmc)
            for dic in self.dic_dlbm:
                if dic == dlbm:
                    col_v = self.dic_dlbm[dic]
                    self.worksheet.write(self.row, col_v, dlmj)
                    self.col_list.append(col_v)
                continue
            self.worksheet.set_row(self.row, 25, cell_format)
        # 隐藏没有数据的列
        for i in range(3, 51, 1):
            if i not in self.col_list:
                self.worksheet.set_column(i, i, 0)
        self.workbook.close()


def get_fields(input_table):
    """
    用于获取表中所有的字段
    """
    desc = GP.Describe(input_table)
    fields = []
    for field in desc.Fields:
        fields.append(field.Name)
    return fields


def overlay():
    """
    三种情况下的分析处理:
        只有DK
        只有PTB
        两者都有
    """
    # intersect_analysis
    if GP.exists(output_xzdw):
        GP.delete_management(output_xzdw)
    if GP.exists(output_dltb):
        GP.delete_management(output_dltb)
    if GP.exists(output_lxdw):
        GP.delete_management(output_lxdw)
    if input_ptb.strip() != "" and input_dk.strip() != "":
        in_tb = input_dltb + ';' + input_dk
        in_xw = input_xzdw + ';' + input_dk
        in_lw = input_lxdw + ';' + input_dk
        fc_management(in_tb, in_xw, in_lw)
    elif input_dk.strip() == "" and input_ptb.strip() != "":
        in_tb = input_dltb + ';' + input_ptb
        in_xw = input_xzdw + ';' + input_ptb
        in_lw = input_lxdw + ';' + input_ptb
        fc_management(in_tb, in_xw, in_lw)
    else:
        # identity_analysis
        in_tb = input_dltb + ';' + input_dk
        in_xw = input_xzdw + ';' + input_dk
        in_lw = input_lxdw + ';' + input_dk
        fc_management(in_tb, in_xw, in_lw)


def fc_management(in_tb, in_xw, in_lw):
    """
	图形数据的叠加分析处理
	"""
    out_dltb = "TDLTB"
    out_xzdw = "TXZDW"
    out_lxdw = "TLXDW"
    if input_dk.strip() != "" and input_ptb.strip() != "":
        GP.AddMessage("Intersecting ...")
        GP.Intersect_analysis(in_tb, out_dltb, "ALL", "", "")
        GP.identity(out_dltb, input_ptb, output_dltb, "ALL", "", "")
        GP.delete_management(out_dltb)
        GP.AddMessage("Identity successful " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    if input_dk.strip() == "" or input_ptb.strip() == "":
        GP.AddMessage("Intersecting ...")
        GP.Intersect_analysis(in_tb, output_dltb, "ALL", "", "")
        GP.AddMessage("Intersect successfully! " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    GP.Intersect_analysis(in_xw, out_xzdw, "ALL", "", "")
    GP.Intersect_analysis(in_lw, out_lxdw, "ALL", "", "")
    update_dltb(output_dltb)
    GP.AddMessage("Update dltb successfully!" + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    GP.identity(out_xzdw, output_dltb, output_xzdw, "ALL", "", "KEEP_RELATIONSHIPS")
    GP.identity(out_lxdw, output_dltb, output_lxdw, "ALL", "", "")
    update_xzdw(output_xzdw)
    GP.AddMessage("Update xzdw successfully! " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    update_lxdw(output_lxdw)
    GP.AddMessage("Update lxdw successfully! " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    GP.delete_management(out_xzdw)
    GP.delete_management(out_lxdw)


def update_dltb(input_table):
    """
	更新DLTB中的TBMJ/TKXS以及面积换算
	"""
    fields = get_fields(input_table)
    if 'BZBGHDLBM' in fields:
        rows = GP.UpdateCursor(input_table)
        row = rows.Next()
        while row:
            if row.BZBGHDLBM.strip() != "":
                row.DLBM = row.BZBGHDLBM
                row.QSXZ = row.BZBGHQSXZ
            if row.TKXS > 1:
                row.TKXS = (row.TKXS/100)
            if is_hectare == "true":
                row.TBMJ = round((row.shape_Area/10000), 4)
            else:
                row.TBMJ = round(row.shape_Area, 2)
            # row.TBMJ = round(row.shape_Area, 2)
            row.TKMJ = round(row.TBMJ*row.TKXS, 2)
            rows.updaterow(row)
            row = rows.Next()
        del row, rows
    else:
        rows = GP.UpdateCursor(input_table)
        row = rows.Next()
        while row:
            if row.TKXS > 1:
                row.TKXS = (row.TKXS/100)
            if is_hectare == "true":
                row.TBMJ = round((row.shape_Area/10000), 4)
                row.TKMJ = round(row.TBMJ*row.TKXS, 4)
            else:
                row.TBMJ = round(row.shape_Area, 2)
                row.TKMJ = round(row.TBMJ*row.TKXS, 2)
            # row.TBMJ = round(row.shape_Area, 2)
            # row.TKMJ = round(row.TBMJ*row.TKXS, 2)
            rows.updaterow(row)
            row = rows.Next()
        del row, rows


def update_xzdw(fc_xzdw):

    fields = get_fields(fc_xzdw)
    rows = GP.UpdateCursor(fc_xzdw)
    row = rows.Next()
    while row:
        # if row.LEFT_BZBGHDLBM.strip() != None:
        row.CD = round(row.SHAPE_LENGTH, 1)
        if is_hectare == "true":
            row.XZDWMJ = round(row.CD*row.KD/10000, 6)
        else:
            row.XZDWMJ = round(row.CD * row.KD, 2)
        # row.XZDWMJ = row.CD * row.KD
        rows.updaterow(row)
        row = rows.Next()
    del row, rows
    if 'LEFT_BZBGHDLBM' in fields or 'RIGHT_BZBGHDLBM' in fields:
        GP.MakeFeatureLayer(fc_xzdw, "temp_xzdw")
        query = '[LEFT_BZBGHDLBM]<>"" AND [RIGHT_BZBGHDLBM]<>"" AND [dlbm] NOT IN("102","118")'
        GP.selectlayerbyattribute("temp_xzdw", "NEW_SELECTION", query)
        GP.deleterows("temp_xzdw")


def update_lxdw(fc_lxdw):

    fields = get_fields(fc_lxdw)
    rows = GP.UpdateCursor(fc_lxdw)
    row = rows.Next()
    while row:
        if is_hectare == "true":
            row.MJ = round(row.MJ/10000, 6)
        rows.updaterow(row)
        row = rows.Next()
    del row, rows
    if 'BZBGHDLBM' in fields:
        GP.MakeFeatureLayer(fc_lxdw, "temp_lxdw")
        query = '[BZBGHDLBM]<>""'
        GP.selectlayerbyattribute("temp_lxdw", "NEW_SELECTION", query)
        GP.deleterows("temp_lxdw")


def create_table():
    t_hzmj = "T_HZMJ"
    t_xzdw = "T_XZDW"
    t_lxdw = "T_LXDW"
    t_dltb = "T_DLTB"
    if GP.exists(t_hzmj):
        GP.delete_management(t_hzmj)
    if GP.exists(t_lxdw):
        GP.delete_management(t_lxdw)
    if GP.exists(t_xzdw):
        GP.delete_management(t_xzdw)
    if GP.exists(t_dltb):
        GP.delete_management(t_dltb)
    if GP.exists("t_hzdlmj"):
        GP.delete_management("t_hzdlmj")
    GP.CreateTable_management(location, t_hzmj)
    # Add Fields To Table T_HZMJ
    GP.AddField(t_hzmj, "ZLDWDM", "TEXT", "", "", "", "ZLDWDM", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField(t_hzmj, "ZLDWMC", "TEXT", "", "", "", "ZLDWMC", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField(t_hzmj, "QSDWDM", "TEXT", "", "", "", "QSDWDM", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField(t_hzmj, "QSDWMC", "TEXT", "", "", "", "QSDWMC", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField(t_hzmj, "QSXZ", "TEXT", "", "", "", "QSXZ", "NULLABLE", "REQUIRED", "#")
    GP.AddField(t_hzmj, "DLBM", "TEXT", "", "", "8", "DLBM", "NULLABLE", "REQUIRED", "#")
    # GP.AddField(t_hzmj,"DLMC","TEXT","","","","DLMC","NULLABLE","REQUIRED","#")
    # GP.AddField(t_hzmj,"TBBH","TEXT","","","","TBBH","NULLABLE","NON_REQUIRED","#")
    # GP.AddField(t_hzmj,"TBMJ","DOUBLE","","","","TBMJ","NULLABLE","NON_REQUIRED","#")
    # GP.AddField(t_hzmj,"LXDWMJ","DOUBLE","","","","LXDWMJ","NULLABLE","NON_REQUIRED","#")
    # GP.AddField(t_hzmj,"XZDWMJ","DOUBLE","","","","XZDWMJ","NULLABLE","NON_REQUIRED","#")
    # GP.AddField(t_hzmj,"TKMJ","DOUBLE","","","#", "TKMJ", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField(t_hzmj, "DLMJ", "DOUBLE", "", "", "#", "DLMJ", "NULLABLE", "NON_REQUIRED", "#")
    GP.AddField(t_hzmj, "TABLENAME", "TEXT", "", "", "", "TABLENAME", "NULLABLE", "REQUIRED", "#")
    GP.AddField(t_hzmj, "STAT_TYPE", "TEXT", "", "", "#", "STAT_TYPE", "NULLABLE", "NON_REQUIRED", "#")


def data_statistic():
    conn = pypyodbc.connect('Driver={Microsoft Access Driver (*.mdb)};DBQ=' + location)
    # conn = pypyodbc.connect('Driver={Microsoft Access Driver (*.mdb)};DBQ=F:/python/ArcGH/2014ptb.mdb')
    cur = conn.cursor()
    cur.execute('''UPDATE PDLTB SET XZDWMJ=0,LXDWMJ=0''')
    cur.commit()
    cur.execute('''update pdltb a,pxzdw b set a.xzdwmj=a.xzdwmj +b.xzdwmj*b.kcbl
    where a.objectid = b.left_pdltb''')
    cur.commit()
    cur.execute('''update pdltb a,pxzdw b set a.xzdwmj=a.xzdwmj +b.xzdwmj*(1-b.kcbl)
    where a.objectid = b.right_pdltb''')
    cur.commit()
    cur.execute('''update pdltb a,plxdw b set a.lxdwmj =a.lxdwmj +b.mj where a.objectid = b.fid_pdltb ''')
    cur.commit()
    cur.execute('''update pdltb set tbdlmj = tbmj-xzdwmj-lxdwmj-tkmj ''')
    cur.commit()
    GP.AddMessage("Update XZDWMJ LXDWMJ TBDLMJ finished " +
                  time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
    # INSERT INTO TABLE T_LXDW
    cur.execute('''select dlbm,zltbbh as tbbh,qsxz,left(zldwdm,12) as tzldwdm,zldwmc,
    left(qsdwdm,12) as tqsdwdm,qsdwmc,mj as lxdwmj,'lxdw' as tablename into t_lxdw from plxdw''')
    # INSERT INTO  TABLE T_XZDW
    cur.execute('''select a.dlbm,a.kctbbh1 as tbbh,a.qsxz,left(a.kctbdwdm1,12) as zldwdm,'' as zldwmc,
    left(a.qsdwdm1,12) as qsdwdm,a.qsdwmc1 as qsdwmc,sum(a.xzdwmj*a.kcbl) as xzdwmj,
    'xzdw' as tablename into t_xzdw from pxzdw a,pdltb b where a.xzdwmj>0 and len(a.kctbdwdm1)>0 and
    len(a.kctbbh1)>0 and a.left_pdltb = b.objectid group by a.dlbm,a.kctbbh1,a.qsxz,left(a.kctbdwdm1,12),
    '',left(a.qsdwdm1,12),a.qsdwmc1''')
    # ADD TO TABLE T_XZDW
    cur.execute('''insert into t_xzdw(dlbm,tbbh,qsxz,zldwdm,zldwmc,qsdwdm,qsdwmc,xzdwmj,tablename)
    select a.dlbm,a.kctbbh2,a.qsxz,left(a.kctbdwdm2,12),'',left(a.qsdwdm2,12),a.qsdwmc2,
    sum(a.xzdwmj-a.xzdwmj*a.kcbl),'xzdw' from pxzdw a,pdltb b where a.xzdwmj>0
    and len(a.kctbdwdm2)>0 and len(a.kctbbh2)>0 and b.objectid = a.right_pdltb group by a.dlbm,
    a.kctbbh2,a.qsxz,left(a.kctbdwdm2,12),'',left(a.qsdwdm2,12),a.qsdwmc2''')
    cur.commit()
    ##############
    # INSERT INTO T_HZMJ
    cur.execute('''select dlbm,tbbh,qsxz,left(zldwdm,12) as tzldwdm,zldwmc,left(qsdwdm,12) as tqsdwdm,qsdwmc,
    sum(tbmj) as ttbmj,sum(tbdlmj) as ttbdlmj,sum(xzdwmj) as xwmj,sum(lxdwmj) as lwmj,sum(tkmj) as ttkmj,
    'dltb' as tablename into t_dltb from pdltb group by dlbm,tbbh,qsxz,left(zldwdm,12),zldwmc,left(qsdwdm,12),qsdwmc''')
    cur.commit()
    GP.AddMessage("Finish insert into t_dltb")
    #
    cur.execute('''insert into t_hzmj(zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,dlmj,tablename) select tzldwdm,zldwmc,tqsdwdm,
    qsdwmc, dlbm,sum(lxdwmj),tablename from t_lxdw group by tzldwdm,zldwmc,tqsdwdm,qsdwmc,dlbm,tablename''')
    cur.commit()
    cur.execute('''insert into t_hzmj(zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,dlmj,tablename) select zldwdm,zldwmc,qsdwdm,
    qsdwmc, dlbm,sum(xzdwmj),tablename from t_xzdw where xzdwmj>0 group by zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,tablename''')
    cur.commit()
    cur.execute('''insert into t_hzmj(zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,dlmj,tablename) select tzldwdm,zldwmc,tqsdwdm,
    qsdwmc, dlbm,sum(ttbdlmj),tablename from t_dltb group by tzldwdm,zldwmc,tqsdwdm,qsdwmc,dlbm,tablename''')
    cur.commit()
    cur.execute('''insert into t_hzmj(zldwdm,zldwmc,qsdwdm,qsdwmc,dlbm,dlmj,tablename) select tzldwdm,zldwmc,tqsdwdm,
    qsdwmc,'123', sum(ttkmj), 'tk' from t_dltb where ttkmj>0 group by tzldwdm,zldwmc,tqsdwdm,qsdwmc''')
    #####
    cur.execute('''select distinct(zldwmc), zldwdm into tt from t_hzmj where zldwmc is not null and zldwmc <>'' ''')
    cur.execute('''update t_hzmj a, tt b set a.zldwmc = b.zldwmc where a.zldwdm = b.zldwdm ''')
    cur.commit()
    #
    cur.execute('''insert into t_hzmj(zldwdm,zldwmc,dlbm,dlmj,stat_type) select zldwdm,zldwmc,dlbm,sum(dlmj),'2' from
    t_hzmj group by zldwdm,zldwmc,dlbm''')
    cur.execute('''insert into t_hzmj(zldwdm,zldwmc,dlbm,dlmj,stat_type) select zldwdm,zldwmc,left(dlbm,2),sum(dlmj),'1'
    from t_hzmj where stat_type = '2' group by zldwdm,zldwmc,left(dlbm,2) ''')
    cur.execute('''insert into t_hzmj(zldwdm,zldwmc,dlbm,dlmj,stat_type) select zldwdm,zldwmc,'zmj',sum(dlmj),'0' from
    t_hzmj where stat_type = '1' group by zldwdm,zldwmc ''')
    cur.execute('''drop table tt''')
    cur.execute('''insert into t_hzmj(dlbm,dlmj) select dlbm,sum(dlmj) from t_hzmj where stat_type = '2'
    group by dlbm ''')
    cur.execute('''insert into t_hzmj(dlbm,dlmj) select dlbm,sum(dlmj) from t_hzmj where stat_type = '1'
    group by dlbm''')
    cur.execute('''insert into t_hzmj(dlbm,dlmj) select 'zmj',sum(dlmj) from t_hzmj  where stat_type ='2' ''')
    cur.commit()
    data = cur.execute('''select zldwdm,zldwmc,dlbm,dlmj from t_hzmj where stat_type in('0','1','2') order by zldwdm ''')
    records = data.fetchall()
    toExcel.write_to_xlsx(records)
    cur.close()
    conn.close()


def main():
    overlay()
    create_table()
    data_statistic()


if __name__ == '__main__':
    input_dltb = "DLTB"
    input_xzdw = "XZDW"
    input_lxdw = "LXDW"
    output_dltb = "PDLTB"
    output_xzdw = "PXZDW"
    output_lxdw = "PLXDW"
    GP = arc.create(9.3)
    # Input feature class
    location = GP.GetParameterAsText(0)
    input_dk = GP.GetParameterAsText(1)
    input_ptb = GP.GetParameterAsText(2)
    is_hectare = GP.GetParameterAsText(3)
    # Set Workspace
    GP.Workspace = location
    toExcel = ManipulateExcel("土地利用现状统计表", location)
    main()
