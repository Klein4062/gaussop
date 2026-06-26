# -*- coding: utf-8 -*-
"""
质量改进诉求流程表单生成器 (v3: 同步审核表定稿)
流程: 提出 → 评审(逐项通过→生成子单) → 改进项分析(接纳→定闭环方法) → 闭环(问题单/需求) → 验收(子项→总项,通过即完成)
填写类型视觉约定: 用户填写(白) / 系统生成(绿,带"系统生成"字样) / 预填可改(紫,带"预填·可改"字样)
输出: 质量改进诉求流程表单.xlsx
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import column_index_from_string as colx, get_column_letter

# ---------------- 样式 ----------------
NAVY="1F3864"; BLUE="2E5496"; STEEL="8EAADB"; LBLUE="D9E1F2"
GRAY="F2F2F2"; CREAM="FFF8E1"; WHITE="FFFFFF"; SUBROW="FBF6E2"
SYS_C="E2EFDA"; PRE_C="E4DFEC"   # 系统/预填 输入区底色

thin=Side(border_style="thin",color="B0B0B0")
BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)

F_TITLE  =Font(name="微软雅黑",size=15,bold=True,color="FFFFFF")
F_SUB    =Font(name="微软雅黑",size=9,italic=True,color="44546A")
F_SECTION=Font(name="微软雅黑",size=11,bold=True,color="FFFFFF")
F_LABEL  =Font(name="微软雅黑",size=10,bold=True,color="333333")
F_VALUE  =Font(name="微软雅黑",size=10,color="000000")
F_HINT   =Font(name="微软雅黑",size=9,italic=True,color="707070")
F_NOTE   =Font(name="微软雅黑",size=9,italic=True,color="808080")
F_BOX    =Font(name="微软雅黑",size=10,bold=True,color="FFFFFF")
F_TBLH   =Font(name="微软雅黑",size=10,bold=True,color="FFFFFF")
F_SAMP   =Font(name="微软雅黑",size=9,color="595959")

FILL_TITLE  =PatternFill("solid",fgColor=NAVY)
FILL_SECTION=PatternFill("solid",fgColor=BLUE)
FILL_TBLH  =PatternFill("solid",fgColor=BLUE)
FILL_BOX    =PatternFill("solid",fgColor=STEEL)
FILL_LABEL  =PatternFill("solid",fgColor=GRAY)
FILL_INPUT  =PatternFill("solid",fgColor=WHITE)
FILL_LONG   =PatternFill("solid",fgColor=CREAM)
FILL_SUB    =PatternFill("solid",fgColor=LBLUE)
FILL_SAMP   =PatternFill("solid",fgColor=SUBROW)
FILL_SYS    =PatternFill("solid",fgColor=SYS_C)
FILL_PRE    =PatternFill("solid",fgColor=PRE_C)

C =Alignment(horizontal="center",vertical="center",wrap_text=True)
L =Alignment(horizontal="left",  vertical="center",wrap_text=True,indent=1)
LT=Alignment(horizontal="left",  vertical="top",   wrap_text=True,indent=1)

# 字段填写类型
USER=None; SYS="sys"; PRE="pre"

def box(ws,r1,c1,r2,c2):
    for r in range(r1,r2+1):
        for c in range(c1,c2+1):
            ws.cell(r,c).border=BORDER

def merge_val(ws,r,c1,c2,value="",fill=FILL_INPUT,font=F_VALUE,align=L):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cell=ws.cell(r,c1,value); cell.fill=fill; cell.font=font; cell.alignment=align
    return cell

def style_input(cell, mode):
    """按填写类型给输入单元格上色与提示"""
    if mode==SYS:
        cell.fill=FILL_SYS; cell.value="(系统生成)"; cell.font=F_HINT; cell.alignment=LT
    elif mode==PRE:
        cell.fill=FILL_PRE; cell.value="(预填·可改)"; cell.font=F_HINT; cell.alignment=LT
    else:
        cell.fill=FILL_INPUT; cell.font=F_VALUE; cell.alignment=LT
    return cell


# ---------------- 枚举 ----------------
OPT={
 "来源":["现网问题","客户提出","实验室发现"],
 "类型":["问题单","需求"],
 "优先级":["高","中","低"],
 "状态":["待评审","评审中","分析中","实施中","待验收","已完成","已驳回"],
 "评审结果":["通过","不通过"],
 "是否":["是","否"],
 "闭环方法":["问题单闭环","需求闭环"],
 "验收是否通过":["通过","不通过"],
}


# ---------------- 表单渲染器 ----------------
def render_form(ws,title,subtitle,sections):
    ws.sheet_view.showGridlines=False
    for c,w in {"A":17,"B":15,"C":15,"D":17,"E":15,"F":15}.items():
        ws.column_dimensions[c].width=w

    def add_dropdown(coord,options):
        if not options: return
        dv=DataValidation(type="list",formula1='"'+",".join(options)+'"',allow_blank=True)
        ws.add_data_validation(dv); dv.add(coord)

    def parse(f):
        # TWO 字段: label | (label,options) | (label,options,mode)
        if not isinstance(f,tuple): return (f,None,USER)
        f=list(f)+[None]*(3-len(f))
        return (f[0], f[1], f[2] or USER)

    r=1
    merge_val(ws,r,1,6,title,FILL_TITLE,F_TITLE,C); ws.row_dimensions[r].height=32; r+=1
    if subtitle:
        merge_val(ws,r,1,6,subtitle,FILL_SUB,F_SUB,C); ws.row_dimensions[r].height=16; r+=1

    for sec in sections:
        merge_val(ws,r,1,6,sec["title"],FILL_SECTION,F_SECTION,L); box(ws,r,1,r,6)
        ws.row_dimensions[r].height=22; r+=1
        for rowdef in sec["fields"]:
            kind=rowdef[0]
            if kind=="SINGLE":           # ("SINGLE", label, height, options, mode)
                label=rowdef[1]; height=rowdef[2] if len(rowdef)>2 else 36
                options=rowdef[3] if len(rowdef)>3 else None
                mode  =rowdef[4] if len(rowdef)>4 else USER
                lc=ws.cell(r,1,label); lc.font=F_LABEL; lc.fill=FILL_LABEL; lc.alignment=C
                cell=merge_val(ws,r,2,6); style_input(cell,mode); box(ws,r,1,r,6)
                if mode!=SYS: add_dropdown(ws.cell(r,2).coordinate,options)
                ws.row_dimensions[r].height=height; r+=1
            elif kind=="TWO":             # ("TWO", f1, f2)
                l1,o1,m1=parse(rowdef[1]); l2,o2,m2=parse(rowdef[2])
                cc=ws.cell(r,1,l1); cc.font=F_LABEL; cc.fill=FILL_LABEL; cc.alignment=C
                c1=merge_val(ws,r,2,3); style_input(c1,m1)
                cc=ws.cell(r,4,l2); cc.font=F_LABEL; cc.fill=FILL_LABEL; cc.alignment=C
                c2=merge_val(ws,r,5,6); style_input(c2,m2); box(ws,r,1,r,6)
                if m1!=SYS: add_dropdown(ws.cell(r,2).coordinate,o1)
                if m2!=SYS: add_dropdown(ws.cell(r,5).coordinate,o2)
                ws.row_dimensions[r].height=24; r+=1
            elif kind=="TABLE":           # ("TABLE", headers, col_spans, n_rows, col_modes, col_opts)
                headers=rowdef[1]; col_spans=rowdef[2]; n_rows=rowdef[3]
                col_modes=rowdef[4] if len(rowdef)>4 else [USER]*len(headers)
                col_opts =rowdef[5] if len(rowdef)>5 else [None]*len(headers)
                for (a,b),h in zip(col_spans,headers):
                    ws.merge_cells(start_row=r,start_column=a,end_row=r,end_column=b)
                    cell=ws.cell(r,a,h); cell.fill=FILL_TBLH; cell.font=F_TBLH; cell.alignment=C
                    box(ws,r,a,r,b)
                ws.row_dimensions[r].height=20; r+=1
                for _ in range(n_rows):
                    for (a,b),mode,opt in zip(col_spans,col_modes,col_opts):
                        ws.merge_cells(start_row=r,start_column=a,end_row=r,end_column=b)
                        cell=ws.cell(r,a); style_input(cell,mode); box(ws,r,a,r,b)
                        if mode!=SYS: add_dropdown(ws.cell(r,a).coordinate,opt)
                    ws.row_dimensions[r].height=26; r+=1
        r+=1
    return r


# =====================================================================
wb=openpyxl.Workbook()

# ---------------- Sheet 0: 流程总览 ----------------
ws=wb.active; ws.title="0-流程总览"; ws.sheet_view.showGridlines=False
for c,w in {"A":17,"B":15,"C":15,"D":17,"E":15,"F":15}.items():
    ws.column_dimensions[c].width=w
merge_val(ws,1,1,6,"质量改进诉求处理流程（提出 → 评审 → 改进项分析 → 闭环 → 验收）",FILL_TITLE,F_TITLE,C)
ws.row_dimensions[1].height=34

# 流程主轴
merge_val(ws,3,1,6,
    "① 提出(含子项)  ➜  ② 评审(逐项通过→生成子单)  ➜  ③ 改进项分析(接纳→定闭环方法)  ➜  ④ 闭环(问题单/需求)  ➜  ⑤ 验收(子项→总项,通过即完成)",
    FILL_BOX,F_BOX,C); box(ws,3,1,3,6); ws.row_dimensions[3].height=34
merge_val(ws,4,1,6,"④ 闭环含两条并行子流程：问题单闭环（缺陷/故障纠正预防）｜ 需求闭环（新功能/优化开发上线）",FILL_SUB,F_NOTE,C)
box(ws,4,1,4,6); ws.row_dimensions[4].height=18

# 规则
merge_val(ws,6,1,6,"单据层级与流转规则",FILL_SECTION,F_SECTION,L); box(ws,6,1,6,6); ws.row_dimensions[6].height=20
rules=[
 "1. 一张【主诉求单】可包含多个改进子项（在提出单的『子项明细表』中逐条列出）。",
 "2. 【评审】对子项逐项判定 通过/不通过；通过则生成【子单】(QI-YYYY-NNN-XX) 进入分析，不通过则该子项终止。",
 "3. 【改进项分析】由分析人决定是否接纳；接纳后决定闭环方法（问题单/需求），并填写闭环单号（校验合法性）后进入闭环。",
 "4. 每张【子单】独立走 闭环 → 子项验收；问题单走问题单闭环单、需求走需求闭环单。",
 "5. 主单下全部【子单】子项验收通过后，进行【总项验收】；总项验收通过即完成（无独立归档环节）。",
]
r=7
for t in rules:
    merge_val(ws,r,1,6,t,FILL_INPUT,F_VALUE,L); box(ws,r,1,r,6); ws.row_dimensions[r].height=26; r+=1

# 阶段说明表
r+=1
hh=["阶段","对应表单(Sheet)","目的","关键判定"]
colspans=[(1,1),(2,2),(3,5),(6,6)]
for (a,b),htxt in zip(colspans,hh):
    ws.merge_cells(start_row=r,start_column=a,end_row=r,end_column=b)
    cell=ws.cell(r,a,htxt); cell.fill=FILL_TBLH; cell.font=F_TBLH; cell.alignment=C
box(ws,r,1,r,6); ws.row_dimensions[r].height=22; r+=1
rows=[
 ("① 提出","1-提出-质量改进诉求单","记录诉求并列出全部子项","是否受理"),
 ("② 评审","2-评审-评审记录单","逐项评审、生成子单","通过/不通过"),
 ("③ 分析","3-分析-改进项分析单","接纳与否、定闭环方法、挂闭环单号","是否接纳；闭环方法"),
 ("④a 闭环","4-1闭环-问题单闭环单","问题单:进展与效果自测","进展/自测"),
 ("④b 闭环","4-2闭环-需求闭环单","需求:进展与效果自测","进展/自测"),
 ("⑤a 验收","5-1-验收-子项验收单","逐子单验收","验收是否通过"),
 ("⑤b 验收","5-2-验收-总项验收单","主单总体验收,通过即完成","总项验收是否通过"),
 ("★ 跟踪","1.5-子单台账(中枢)","汇总全部子单各环节进度","主单是否具备总项验收条件"),
]
for row in rows:
    for (a,b),val in zip(colspans,row):
        ws.merge_cells(start_row=r,start_column=a,end_row=r,end_column=b)
        cell=ws.cell(r,a,val); cell.font=F_VALUE; cell.alignment=LT
        if a==1: cell.fill=FILL_LABEL; cell.font=F_LABEL; cell.alignment=C
    box(ws,r,1,r,6); ws.row_dimensions[r].height=30; r+=1


# ---------------- Sheet 1: 提出 ----------------
ws1=wb.create_sheet("1-提出-质量改进诉求单")
render_form(ws1,"质量改进诉求单（提出阶段 · 可含多个子项）",
    "主单号 QI-YYYY-NNN    ①提出 → ②评审 → ③改进项分析 → ④闭环 → ⑤验收",[
    {"title":"一、主单信息","fields":[
        ("TWO",("主诉求单号",None,SYS),("提出日期",None,SYS)),
        ("SINGLE","当前状态",22,OPT["状态"],SYS),
    ]},
    {"title":"二、人员信息","fields":[
        ("SINGLE","提出人(工号+姓名)",24,None,SYS),
    ]},
    {"title":"三、诉求总体内容","fields":[
        ("SINGLE","诉求总体标题",26),
        ("SINGLE","诉求来源",22,OPT["来源"]),
        ("SINGLE","总体背景与目标（问题概述、影响、期望达成）",50),
    ]},
    {"title":"四、改进子项明细（逐行一个子项；评审后将按子项生成子单）","fields":[
        ("TABLE",
         ["子项编号","子项标题","类型","子项简述/现象","优先级"],
         [(1,1),(2,3),(4,4),(5,5),(6,6)],6,
         [SYS,USER,USER,USER,USER],
         [None,None,OPT["类型"],None,OPT["优先级"]]),
    ]},
    {"title":"五、审批","fields":[
        ("SINGLE","审批人(工号+姓名)",24),
    ]},
])


# ---------------- Sheet 2: 评审 ----------------
ws2=wb.create_sheet("2-评审-评审记录单")
render_form(ws2,"评审记录单（评审阶段 · 逐项评审,通过则生成子单）",
    "关联主诉求单号    评审通过则生成子单进入分析,不通过则该子项终止",[
    {"title":"一、评审信息","fields":[
        ("TWO",("评审单号",None,SYS),("关联主诉求单号",None,SYS)),
        ("TWO",("评审日期",None,SYS),("评审人(工号+姓名)",None,SYS)),
    ]},
    {"title":"二、子项拆单分发表（逐项评审；子项内容继承自提出单,此处只判定通过与否）","fields":[
        ("TABLE",
         ["子项编号","评审结果","不通过理由","生成子单号","责任人/部门"],
         [(1,1),(2,2),(3,4),(5,5),(6,6)],6,
         [SYS,USER,USER,SYS,USER],
         [None,OPT["评审结果"],None,None,None]),
    ]},
    {"title":"三、统计","fields":[
        ("SINGLE","拆单总数(张)",22,None,SYS),
    ]},
])


# ---------------- Sheet 1.5: 子单台账（中枢） ----------------
wsL=wb.create_sheet("1.5-子单台账(中枢)")
wsL.sheet_view.showGridlines=False
heads=["子单号","父诉求单号","子项编号","类型","责任人/部门","评审结果",
       "分析·是否接纳","闭环方法","闭环单号","子项验收·是否通过","当前状态"]
widths=[16,15,9,9,14,11,12,12,13,15,11]
for i,w in enumerate(widths,start=1):
    wsL.column_dimensions[get_column_letter(i)].width=w
merge_val(wsL,1,1,11,"子单台账（中枢 · 汇总全部子单在各环节的进度）",FILL_TITLE,F_TITLE,C)
wsL.row_dimensions[1].height=32
merge_val(wsL,2,1,11,"一张主诉求单拆为多张子单,每张子单独立流转;本表以子单为单位汇总各阶段结果与状态。"
                     "全部子单『子项验收·是否通过』为通过后,主单可进行总项验收。",FILL_SUB,F_SUB,L)
wsL.row_dimensions[2].height=34
r=3
for i,h in enumerate(heads,start=1):
    cell=wsL.cell(r,i,h); cell.fill=FILL_TBLH; cell.font=F_TBLH; cell.alignment=C; cell.border=BORDER
wsL.row_dimensions[r].height=26
samples=[
 ["QI-2026-001-01","QI-2026-001","01","问题单","张工/工艺部","通过","是","问题单闭环","PC-2026-001","通过","已完成"],
 ["QI-2026-001-02","QI-2026-001","02","需求","李工/研发部","通过","是","需求闭环","RC-2026-001","待验收","待验收"],
 ["QI-2026-001-03","QI-2026-001","03","问题单","王工/质量部","不通过","—","—","—","—","已驳回"],
]
r=4
for row in samples:
    for i,v in enumerate(row,start=1):
        cell=wsL.cell(r,i,v); cell.fill=FILL_SAMP; cell.font=F_SAMP
        cell.alignment=C if i!=5 else LT; cell.border=BORDER
    wsL.row_dimensions[r].height=22; r+=1
for _ in range(8):
    for i in range(1,12):
        cell=wsL.cell(r,i); cell.fill=FILL_INPUT; cell.border=BORDER; cell.alignment=LT
    wsL.row_dimensions[r].height=22; r+=1
merge_val(wsL,r+1,1,11,"注：浅色行为示例数据,实际使用时替换或删除。子单号 = 父诉求单号 + 子项序号。",FILL_SUB,F_NOTE,L)


# ---------------- Sheet 3: 改进项分析 ----------------
ws3=wb.create_sheet("3-分析-改进项分析单")
render_form(ws3,"改进项分析单（分析阶段 · 针对单张子单）",
    "关联：子单号 + 父诉求单号    接纳后填写闭环单号(校验合法性)方可进入闭环;不接纳则终止",[
    {"title":"一、单据信息","fields":[
        ("TWO",("分析单号",None,SYS),("关联子单号",None,SYS)),
        ("TWO",("父诉求单号",None,SYS),("分析日期",None,SYS)),
        ("SINGLE","当前状态",22,OPT["状态"],SYS),
    ]},
    {"title":"二、分析与接纳","fields":[
        ("TWO","接纳版本",("是否接纳",OPT["是否"])),
        ("TWO",("闭环方法",OPT["闭环方法"]),"计划完成期限"),
        ("SINGLE","不接纳理由（是否接纳为『否』时必填）",40),
    ]},
    {"title":"三、闭环单号（必填,需校验合法性）","fields":[
        ("SINGLE","问题单号/需求单号",26),
    ]},
    {"title":"四、分析进展子项（可逐条单独提交,作为外部审视进展的依据）","fields":[
        ("TABLE",
         ["序号","进展说明","提交时间","提交人"],
         [(1,1),(2,4),(5,5),(6,6)],6,
         [SYS,USER,SYS,SYS],
         [None,None,None,None]),
    ]},
])


# ---------------- Sheet 4-1: 问题单闭环 ----------------
ws4a=wb.create_sheet("4-1闭环-问题单闭环单")
render_form(ws4a,"问题单闭环单（闭环 · 问题单路径 · 针对单张子单）",
    "关联：子单号 + 父诉求单号 + 分析单号    只关注当前进展与闭环效果自测",[
    {"title":"一、单据信息","fields":[
        ("TWO",("问题单号",None,SYS),("关联子单号",None,SYS)),
        ("TWO",("父诉求单号",None,SYS),("关联分析单号",None,SYS)),
        ("TWO",("闭环类型",None,SYS),("当前状态",OPT["状态"],SYS)),
    ]},
    {"title":"二、进展与自测","fields":[
        ("SINGLE","当前进展",50),
        ("SINGLE","闭环效果自测",50),
    ]},
    {"title":"三、进展子项（可逐条单独提交,作为外部审视进展的依据）","fields":[
        ("TABLE",
         ["序号","进展说明","提交时间","提交人"],
         [(1,1),(2,4),(5,5),(6,6)],6,
         [SYS,USER,SYS,SYS],
         [None,None,None,None]),
    ]},
])


# ---------------- Sheet 4-2: 需求闭环 ----------------
ws4b=wb.create_sheet("4-2闭环-需求闭环单")
render_form(ws4b,"需求闭环单（闭环 · 需求路径 · 针对单张子单）",
    "关联：子单号 + 父诉求单号 + 分析单号    只关注当前进展与闭环效果自测",[
    {"title":"一、单据信息","fields":[
        ("TWO",("需求单号",None,SYS),("关联子单号",None,SYS)),
        ("TWO",("父诉求单号",None,SYS),("关联分析单号",None,SYS)),
        ("TWO",("闭环类型",None,SYS),("当前状态",OPT["状态"],SYS)),
    ]},
    {"title":"二、进展与自测","fields":[
        ("SINGLE","当前进展",50),
        ("SINGLE","闭环效果自测",50),
    ]},
    {"title":"三、进展子项（可逐条单独提交,作为外部审视进展的依据）","fields":[
        ("TABLE",
         ["序号","进展说明","提交时间","提交人"],
         [(1,1),(2,4),(5,5),(6,6)],6,
         [SYS,USER,SYS,SYS],
         [None,None,None,None]),
    ]},
])


# ---------------- Sheet 5-1: 子项验收 ----------------
ws5a=wb.create_sheet("5-1-验收-子项验收单")
render_form(ws5a,"子项验收单（针对单张子单;全部子单通过后进入总项验收）",
    "关联：子单号 + 父诉求单号 + 闭环单号",[
    {"title":"一、单据信息","fields":[
        ("TWO",("验收单号",None,SYS),("关联子单号",None,SYS)),
        ("TWO",("父诉求单号",None,SYS),("关联闭环单号",None,SYS)),
        ("TWO",("验收人(工号+姓名)",None,PRE),("当前状态",OPT["状态"],SYS)),
    ]},
    {"title":"二、验收结论","fields":[
        ("SINGLE","验收结论",50),
        ("SINGLE","验收是否通过",24,OPT["验收是否通过"]),
    ]},
])


# ---------------- Sheet 5-2: 总项验收 ----------------
ws5b=wb.create_sheet("5-2-验收-总项验收单")
render_form(ws5b,"总项验收单（主单下全部子项验收通过后进行,通过即完成）",
    "关联：主诉求单号",[
    {"title":"一、单据信息","fields":[
        ("TWO",("总项验收单号",None,SYS),("关联主诉求单号",None,SYS)),
        ("TWO",("验收人(工号+姓名)",None,PRE),("当前状态",OPT["状态"],SYS)),
    ]},
    {"title":"二、子项验收状态（系统生成,只读）","fields":[
        ("SINGLE","各子项验收状态",50,None,SYS),
    ]},
    {"title":"三、总项验收结论","fields":[
        ("SINGLE","总项验收结论",50),
        ("SINGLE","总项验收是否通过",24,OPT["验收是否通过"]),
    ]},
])


# ---------------- Sheet 6: 填写约定与编号规则 ----------------
ws6=wb.create_sheet("6-填写约定与编号规则")
ws6.sheet_view.showGridlines=False
for col,w in {"A":22,"B":16,"C":64}.items():
    ws6.column_dimensions[col].width=w
merge_val(ws6,1,1,3,"填写约定与编号规则",FILL_TITLE,F_TITLE,C)
ws6.row_dimensions[1].height=30

# 填写类型图例
merge_val(ws6,2,1,3,"一、填写类型（输入区底色）",FILL_SECTION,F_SECTION,L); box(ws6,2,1,2,3); ws6.row_dimensions[2].height=20
legend=[
 ("用户填写","白底","由用户录入(必填/选填/条件必填)"),
 ("系统生成","绿底(带『系统生成』)","系统自动生成/继承,用户免填(编号/状态/关联/日期/统计等)"),
 ("预填可改","紫底(带『预填·可改』)","系统预填默认值(如验收人=提出人),用户可修改"),
]
r=3
for a,b,cc in legend:
    ws6.cell(r,1,a).font=F_LABEL
    cell=ws6.cell(r,2,b)
    cell.fill={"白底":FILL_INPUT,"绿底(带『系统生成』)":FILL_SYS,"紫底(带『预填·可改』)":FILL_PRE}[b]
    cell.font=F_HINT if b!="白底" else F_VALUE; cell.alignment=C
    ws6.cell(r,3,cc).font=F_VALUE
    for i in range(1,4): ws6.cell(r,i).border=BORDER; ws6.cell(r,i).alignment=L if i==3 else C
    ws6.row_dimensions[r].height=24; r+=1

# 编号规则
r+=1
merge_val(ws6,r,1,3,"二、编号规则",FILL_SECTION,F_SECTION,L); box(ws6,r,1,r,3); ws6.row_dimensions[r].height=20; r+=1
for i,h in enumerate(["编号","格式","说明"],start=1):
    cell=ws6.cell(r,i,h); cell.fill=FILL_TBLH; cell.font=F_TBLH; cell.alignment=C; cell.border=BORDER
ws6.row_dimensions[r].height=22; r+=1
ids=[
 ("主诉求单号","QI-YYYY-NNN","主单唯一编号,全流程主键"),
 ("子项编号","01 / 02 / 03…","主单内子项序号"),
 ("子单号","QI-YYYY-NNN-XX","= 主诉求单号 + 子项序号;评审通过时生成"),
 ("分析单号","AN-YYYY-NNN","改进项分析单编号(系统生成)"),
 ("问题单号","PC-YYYY-NNN","问题单闭环单编号(系统生成)"),
 ("需求单号","RC-YYYY-NNN","需求闭环单编号(系统生成)"),
 ("子项验收单号","AC-YYYY-NNN","子项验收单编号(系统生成)"),
 ("总项验收单号","系统生成,关联主单","总项验收单编号(系统生成)"),
]
for a,b,cc in ids:
    ws6.cell(r,1,a).font=F_LABEL; ws6.cell(r,1,a).fill=FILL_LABEL
    ws6.cell(r,2,b).font=F_VALUE; ws6.cell(r,3,cc).font=F_VALUE
    for i in range(1,4): ws6.cell(r,i).border=BORDER; ws6.cell(r,i).alignment=C if i==2 else L
    ws6.row_dimensions[r].height=24; r+=1
r+=1
merge_val(ws6,r,1,3,"三、流程：提出 → 评审(逐项通过→生成子单) → 改进项分析(接纳→定闭环方法) → 闭环(问题单/需求) → 验收(子项→总项,通过即完成)",
          FILL_SUB,F_NOTE,L); ws6.row_dimensions[r].height=40


# ---------------- 调整 Sheet 顺序 ----------------
order=["0-流程总览","1-提出-质量改进诉求单","1.5-子单台账(中枢)","2-评审-评审记录单",
       "3-分析-改进项分析单","4-1闭环-问题单闭环单","4-2闭环-需求闭环单",
       "5-1-验收-子项验收单","5-2-验收-总项验收单","6-填写约定与编号规则"]
wb._sheets.sort(key=lambda s: order.index(s.title))

# ---------------- 保存 ----------------
out=os.path.join(os.path.dirname(os.path.abspath(__file__)),"质量改进诉求流程表单.xlsx")
wb.save(out)
print("saved:",out)
print("sheets:",wb.sheetnames)
