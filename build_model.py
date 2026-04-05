#!/usr/bin/env python3
"""
Three-Way Corporate Finance Model — Electricity Grid Transmission Utility
Australian AER Regulatory Framework: PTRM | REFM | Frontier | RAB | DTM
Period: FY2019A-FY2023A (Historical) | FY2024E-FY2033E (Forecast)
All amounts: AUD Millions unless stated
"""
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime

# ── Year config ────────────────────────────────────
HIST = list(range(2019, 2024))
FCST = list(range(2024, 2034))
ALL  = HIST + FCST
NH, NF, NT = 5, 10, 15
LC = 1; DC = 2  # label col, data start col

def ycol(yr): return DC + yr - 2019
def ylet(yr): return get_column_letter(ycol(yr))

# ── Styles ─────────────────────────────────────────
def fill(h): return PatternFill("solid", fgColor=h)

C_HDR = fill("003366"); C_SEC = fill("D9E1F2"); C_INP = fill("E2EFDA")
C_HIS = fill("F2F2F2"); C_SUB = fill("FFF2CC"); C_TOT = fill("D9E1F2")
C_WHT = fill("FFFFFF"); C_COV = fill("001F4D"); C_OK  = fill("C6EFCE")
C_ERR = fill("FFC7CE")

def mkf(bold=False, italic=False, color="000000", size=10):
    return Font(name="Calibri", size=size, bold=bold, italic=italic, color=color)
F_H   = mkf(bold=True, color="FFFFFF", size=11)
F_S   = mkf(bold=True, color="003366")
F_L   = mkf()
F_I   = mkf(color="1B6B1B")
F_HV  = mkf(italic=True, color="595959")
F_C   = mkf()
F_B   = mkf(bold=True)
F_YHH = mkf(bold=True, color="595959")
F_YHF = mkf(bold=True, color="003366")

FMTD = '#,##0;(#,##0)'
FMTP = '0.0%'; FMTP2 = '0.00%'; FMTN = '#,##0'
RA = Alignment(horizontal='right',  vertical='center')
LA = Alignment(horizontal='left',   vertical='center')
CA = Alignment(horizontal='center', vertical='center')

# ── Row maps (pre-defined for cross-sheet formula building) ──
A = {}  # Assumptions
RF = {} # REFM
FR = {} # Frontier
RB = {} # RAB
DT = {} # DTM
PT = {} # PTRM
IS = {} # Income Statement
BS = {} # Balance Sheet
CF = {} # Cash Flow

A.update({'title':1,'yr_hdr':2,'wacc_sec':4,'wacc':5,'rfr':6,'erp':7,'beta':8,
    'gearing':9,'equity_pct':10,'gamma':11,'tax_rate':12,
    'cpi_sec':14,'cpi':15,'vol_sec':17,'ghw':18,'tariff':19,'conn_fees':20,
    'opex_sec':22,'om':23,'corp':24,'insur':25,'reg':26,
    'capex_sec':28,'aug':29,'repl':30,'conn':31,
    'rab_sec':33,'open_rab':34,'std_life':35,'open_reg_dep':36,
    'ppe_sec':38,'open_ppe':39,'open_acc_da':40,'acc_life':41,
    'fin_sec':43,'open_debt':44,'open_cash':45,'cod':46,'div_pay':47,
    'debt_iss':48,'debt_rep':49,'wc_sec':51,'dso':52,'dpo':53,
    'front_sec':55,'front_shift':56,'opex_eff':57,'capex_eff':58,
    'hist_opex_sec':60,'h_om':61,'h_corp':62,'h_insur':63,'h_reg':64,
    'hist_capex_sec':66,'h_aug':67,'h_repl':68,'h_conn':69,
    'hist_fin_sec':71,'h_iss':72,'h_rep':73,
    'hist_bs_sec':75,'open_re':76,'open_cap':77})

RF.update({'title':1,'yr_hdr':2,'opex_sec':4,'om':5,'corp':6,'insur':7,'reg':8,
    'tot_opex':9,'capex_sec':11,'aug':12,'repl':13,'conn':14,'tot_capex':15,
    'var_sec':17,'opex_allowed':18,'opex_actual':19,'opex_var':20,
    'ecm_sec':22,'ecm_rate':23,'ecm_ben':24})

FR.update({'title':1,'yr_hdr':2,'sec':4,'cum_shift':5,'raw_opex':6,'adj_opex':7,
    'raw_capex':8,'adj_capex':9,'xfactor':11})

RB.update({'title':1,'yr_hdr':2,'sec':4,'open_rab':5,'cpi_idx':6,'capex':7,
    'reg_dep':8,'close_rab':9,'avg_rab':11,'accum_sec':13,'accum_reg_dep':14})

DT.update({'title':1,'hdr':2,'instr':3,'rfr_sec':5,'rfr':6,'drp_sec':8,'drp':9,
    'yield_sec':11,'bench_yield':12,'iss_cost':13,'total_yield':14,
    'trail_sec':16,'trail_avg':17,'n_vintages':18,'model_sec':20,'model_wacd':21})

PT.update({'title':1,'yr_hdr':2,'sec':4,'ret_on':5,'ret_of':6,'opex_all':7,
    'tax_all':8,'mar':9,'tax_sec':11,'taxable':12,'tax_gross':13,
    'gamma_ben':14,'net_tax':15})

IS.update({'title':1,'yr_hdr':2,'rev_sec':4,'trans_rev':5,'conn_rev':6,'tot_rev':7,
    'opex_sec':9,'om':10,'corp':11,'insur':12,'reg':13,'tot_opex':14,
    'ebitda':16,'da':18,'ebit':20,'int_exp':22,'ebt':24,'tax':26,
    'net_inc':28,'divs':29,'retd':30})

BS.update({'title':1,'yr_hdr':2,'asset_sec':4,'cash':5,'ar':6,'tot_ca':7,
    'gross_ppe':9,'acc_da':10,'net_ppe':11,'reg_asset':12,'tot_nca':13,
    'tot_assets':14,'liab_sec':16,'ap':17,'tot_cl':18,'lt_debt':20,
    'def_tax':21,'tot_ncl':22,'tot_liab':23,'eq_sec':25,'sh_cap':26,
    'ret_earn':27,'tot_eq':28,'tot_le':29})

CF.update({'title':1,'yr_hdr':2,'ocf_sec':4,'net_inc':5,'da_add':6,'d_ar':7,
    'd_ap':8,'d_reg':9,'tot_ocf':10,'icf_sec':12,'capex':13,'tot_icf':14,
    'fcf_sec':16,'debt_iss':17,'debt_rep':18,'divs_paid':19,'tot_fcf':20,
    'net_cf':22,'open_cash':23,'end_cash':24})


# ── Input data ─────────────────────────────────────
AV = {'wacc':0.065,'rfr':0.043,'erp':0.062,'beta':0.45,'gearing':0.55,
    'gamma':0.585,'tax_rate':0.30,'std_life':40,'acc_life':35,
    'open_rab':4800,'open_reg_dep':1200,'open_ppe':7800,'open_acc_da':2600,
    'open_debt':3500,'open_cash':120,'cod':0.055,'div_pay':0.70,
    'dso':45,'dpo':30,'front_shift':0.005,'opex_eff':0.95,'capex_eff':0.95,
    'open_re':420,'open_cap':880}

YV = {'cpi':[0.025]*10,
    'om':[100,102,104,106,108,110,112,114,116,118],
    'corp':[58,59,60,61,62,63,64,65,66,67],
    'insur':[20,20,21,21,22,22,23,23,24,24],
    'reg':[10,10,10,10,10,11,11,11,11,11],
    'aug':[200,210,220,230,240,250,260,270,280,290],
    'repl':[160,165,170,175,180,185,190,195,200,205],
    'conn':[45,47,48,50,52,54,56,58,60,62],
    'debt_iss':[350,360,370,380,390,400,410,420,430,440],
    'debt_rep':[280,285,290,295,300,305,310,315,320,325],
    'ghw':[52000,53000,54000,55000,56000,57000,58000,59000,60000,61000],
    'tariff':[13.2,13.5,13.8,14.1,14.4,14.7,15.0,15.3,15.6,15.9],
    'conn_fees':[8.0,8.2,8.4,8.6,8.8,9.0,9.2,9.4,9.6,9.8]}

HV = {'cpi':[0.015,0.017,0.018,0.030,0.060],
    'om':[89,92,95,97,100],'corp':[52,53,54,56,57],
    'insur':[18,18,18,19,19],'reg':[9,9,9,9,9],
    'aug':[150,165,175,185,195],'repl':[100,105,115,130,145],
    'conn':[30,35,40,40,40],
    'debt_iss':[300,310,320,340,360],'debt_rep':[240,245,250,260,270]}

H_REV  = [603,629,666,687,731]
H_CONN = [6.0,6.5,7.0,7.5,8.0]

# DTM vintage data (FY2014-FY2033)
DTM_YRS = list(range(2014,2034))
DTM_RFR = {2014:3.50,2015:3.00,2016:2.50,2017:2.70,2018:2.90,
    2019:1.70,2020:0.90,2021:1.80,2022:3.70,2023:4.30,
    2024:4.00,2025:3.80,2026:3.70,2027:3.60,2028:3.60,
    2029:3.60,2030:3.60,2031:3.60,2032:3.60,2033:3.60}
DTM_DRP = {2014:1.70,2015:1.80,2016:2.00,2017:1.90,2018:1.80,
    2019:2.00,2020:2.50,2021:2.20,2022:2.00,2023:1.90,
    2024:1.80,2025:1.75,2026:1.70,2027:1.70,2028:1.70,
    2029:1.70,2030:1.70,2031:1.70,2032:1.70,2033:1.70}
ISS_COST = 0.00125  # 0.125% issuance cost

# Pre-compute consistent historical financials
def _calc_hist():
    ppe=AV['open_ppe']; acc_da=AV['open_acc_da']
    debt=AV['open_debt']; re=AV['open_re']; cash=AV['open_cash']
    rab=AV['open_rab']; acc_rd=AV['open_reg_dep']
    R={}
    for idx,yr in enumerate(HIST):
        cap=HV['aug'][idx]+HV['repl'][idx]+HV['conn'][idx]
        ppe += cap
        da  = round(ppe/AV['acc_life'],1)
        acc_da += da
        net_ppe = ppe-acc_da
        cpi_h=HV['cpi'][idx]
        cpi_idx_v=round(rab*cpi_h,1)
        reg_dep=round(rab/AV['std_life'],1)
        rab_close=round(rab+cpi_idx_v+cap-reg_dep,1)
        acc_rd += reg_dep
        avg_rab=round((rab+rab_close)/2,1)
        rab=rab_close
        rev=H_REV[idx]; conn=H_CONN[idx]; trans=rev-conn
        tot_op=sum(HV[k][idx] for k in ['om','corp','insur','reg'])
        ebitda=rev-tot_op; ebit=round(ebitda-da,1)
        int_e=round(debt*AV['cod'],1)
        ebt=round(ebit-int_e,1)
        tax=round(max(ebt*AV['tax_rate'],0),1)
        ni=round(ebt-tax,1); divs=round(ni*AV['div_pay'],1); retd=round(ni-divs,1)
        re += retd
        iss=HV['debt_iss'][idx]; rep=HV['debt_rep'][idx]
        debt += iss-rep
        ar=round(rev*AV['dso']/365,1)
        ap=round(tot_op*AV['dpo']/365,1)
        reg_asset=round(rab_close-net_ppe,1)
        def_tax=round(max((acc_rd-acc_da)*AV['tax_rate'],0),1)
        tot_assets_v=cash+ar+net_ppe+reg_asset
        tot_le_v=ap+debt+def_tax+AV['open_cap']+re
        # adjust cash to balance
        cash=round(tot_le_v-(ar+net_ppe+reg_asset),1)
        # CFS
        prev_ar = round(H_REV[idx-1]*AV['dso']/365,1) if idx>0 else round(H_REV[0]*0.95*AV['dso']/365,1)
        prev_ap = round(sum(HV[k][idx-1] for k in ['om','corp','insur','reg'])*AV['dpo']/365,1) if idx>0 else ap*0.95
        d_ar=-(ar-prev_ar); d_ap=ap-prev_ap
        ocf=round(ni+da+d_ar+d_ap,1)
        icf=-cap
        fcf=round(iss-rep-divs,1)
        net_cf=round(ocf+icf+fcf,1)
        open_c=R[idx-1]['end_cash'] if idx>0 else AV['open_cash']
        end_c=round(open_c+net_cf,1)
        cash=end_c  # use CFS-derived cash for BS
        R[idx]={'yr':yr,'ppe':ppe,'acc_da':acc_da,'net_ppe':net_ppe,
            'rab_close':rab_close,'avg_rab':avg_rab,'acc_rd':acc_rd,
            'reg_asset':reg_asset,'def_tax':def_tax,
            'lt_debt':debt,'re':re,'ar':ar,'ap':ap,'cash':cash,
            'rev':rev,'conn':conn,'trans':trans,'tot_op':tot_op,
            'da':da,'ebitda':ebitda,'ebit':ebit,'int_e':int_e,
            'ebt':ebt,'tax':tax,'ni':ni,'divs':divs,'retd':retd,
            'iss':iss,'rep':rep,'ocf':ocf,'icf':icf,'fcf':fcf,
            'net_cf':net_cf,'open_cash':open_c,'end_cash':end_c,
            'd_ar':d_ar,'d_ap':d_ap,'reg_dep':reg_dep,'cpi_idx_v':cpi_idx_v,'cap':cap}
    return R

HIST_DATA = _calc_hist()


# ── Helper functions ───────────────────────────────
def xref(sheet, cl, rw):
    q="'" if any(c in sheet for c in " ()") else ""
    return f"={q}{sheet}{q}!{cl}{rw}"

def col_wid(ws, lw=42, dw=11):
    ws.column_dimensions['A'].width=lw
    for i in range(NT):
        ws.column_dimensions[get_column_letter(DC+i)].width=dw

def frz(ws): ws.freeze_panes=ws.cell(row=3, column=DC)

def yr_hdr(ws, row, label="$AUD Millions"):
    ws.row_dimensions[row].height=20
    c=ws.cell(row=row,column=LC,value=label)
    c.fill=C_HDR; c.font=F_H; c.alignment=LA
    for i,yr in enumerate(ALL):
        col=DC+i; suf="A" if yr<2024 else "E"
        c=ws.cell(row=row,column=col,value=f"FY{yr}{suf}")
        c.fill=C_HDR; c.font=F_YHH if yr<2024 else F_YHF; c.alignment=CA

def sec(ws, row, label, ncols=None):
    ws.row_dimensions[row].height=16
    ws.cell(row=row,column=1,value=label).font=F_S
    nc=ncols or (DC+NT-1)
    for c in range(1,nc+2): ws.cell(row=row,column=c).fill=C_SEC

def lbl(ws, row, text, indent=0):
    c=ws.cell(row=row,column=1,value=("    "*indent)+text)
    c.font=F_L; c.alignment=LA

def inp(ws, row, col, val, fmt=FMTD):
    c=ws.cell(row=row,column=col,value=val)
    c.number_format=fmt; c.fill=C_INP; c.font=F_I; c.alignment=RA; return c

def hcell(ws, row, col, val, fmt=FMTD, bold=False):
    c=ws.cell(row=row,column=col,value=val)
    c.number_format=fmt; c.fill=C_HIS
    c.font=mkf(italic=True,color="595959",bold=bold); c.alignment=RA; return c

def fcell(ws, row, col, expr, fmt=FMTD, bold=False, sub=False, tot=False, hist=False):
    c=ws.cell(row=row,column=col,value=expr)
    c.number_format=fmt
    if tot:    c.fill=C_TOT; c.font=F_B
    elif sub:  c.fill=C_SUB; c.font=F_B
    elif hist: c.fill=C_HIS; c.font=mkf(italic=True,color="595959",bold=bold)
    else:      c.fill=C_WHT; c.font=F_B if bold else F_C
    c.alignment=RA; return c

def title_row(ws, text, color="003366"):
    ws.merge_cells(f"A1:{get_column_letter(DC+NT-1)}1")
    c=ws.cell(row=1,column=1,value=text)
    c.fill=fill(color); c.font=mkf(bold=True,color="FFFFFF",size=13)
    c.alignment=LA; ws.row_dimensions[1].height=26


# ── Cover ──────────────────────────────────────────
def build_cover(ws):
    ws.sheet_view.showGridLines=False
    ws.column_dimensions['A'].width=2
    ws.column_dimensions['B'].width=55
    ws.column_dimensions['C'].width=32
    for r in range(1,52):
        for c in range(1,18): ws.cell(row=r,column=c).fill=C_COV
    def cw(r,col,val,fnt=None):
        c=ws.cell(row=r,column=col,value=val)
        c.font=fnt or mkf(color="FFFFFF",size=12); c.alignment=LA; return c
    ws.row_dimensions[4].height=42
    cw(4,2,"ELECTRICITY GRID TRANSMISSION",mkf(bold=True,color="FFFFFF",size=22))
    ws.row_dimensions[5].height=36
    cw(5,2,"THREE-WAY CORPORATE FINANCE MODEL",mkf(bold=True,color="FFFFFF",size=20))
    cw(7,2,"Australian Energy Regulator (AER) | PTRM | REFM | Frontier | RAB | DTM",
       mkf(color="AAD4F5",size=13))
    for c in range(2,14): ws.cell(row=9,column=c).fill=fill("FFD700")
    ws.row_dimensions[9].height=3
    info=[("Company","GridCo Transmission Pty Ltd"),("ABN","12 345 678 901"),
        ("Version","v1.0  |  AER DTM Integrated"),
        ("Date",datetime.date.today().strftime("%d %B %Y")),
        ("Scope","FY2019A-FY2023A Historical  |  FY2024E-FY2033E Forecast"),
        ("Reg Periods","RCP1: FY2024-FY2028  |  RCP2: FY2029-FY2033"),
        ("Jurisdiction","Australia — AER Regulated Transmission"),
        ("Currency","AUD Millions (unless stated)")]
    for i,(k,v) in enumerate(info):
        r=12+i*2; ws.row_dimensions[r].height=20
        cw(r,2,k,mkf(bold=True,color="AAD4F5",size=11))
        cw(r,3,v,mkf(color="FFFFFF",size=11))
    cw(32,2,"MODEL NAVIGATION",mkf(bold=True,color="FFD700",size=12))
    tabs=[("Assumptions","All inputs — the ONLY cells to change"),
        ("DTM","AER Debt Term Model: 10yr trailing avg cost of debt"),
        ("REFM","Revenue & Expenditure Forecasting Model (opex/capex)"),
        ("Frontier","Frontier efficiency shift and X-factor"),
        ("RAB","Regulatory Asset Base roll-forward"),
        ("PTRM","Post Tax Revenue Model — Maximum Allowed Revenue (MAR)"),
        ("Income Statement","5yr historical + 10yr forecast P&L"),
        ("Balance Sheet","5yr historical + 10yr forecast Balance Sheet"),
        ("Cash Flow","5yr historical + 10yr forecast Cash Flow"),
        ("Checks","Model integrity: BS, CFS & RAB checks")]
    for i,(t,d) in enumerate(tabs):
        r=34+i
        cw(r,2,t,mkf(bold=True,color="FFD700",size=10))
        cw(r,3,d,mkf(color="CCCCCC",size=10))


# ── Assumptions ────────────────────────────────────
def build_assumptions(ws):
    ws.sheet_view.showGridLines=False
    col_wid(ws,42,12); frz(ws)
    title_row(ws,"ASSUMPTIONS — Electricity Grid Transmission Utility (AER Framework)")
    yr_hdr(ws, A['yr_hdr'])
    sec(ws,A['wacc_sec'],"WACC & Regulatory Parameters")
    for row,lbl_t,key,fmt in [
        (A['wacc'],"Post-Tax Nominal WACC (%)","wacc",FMTP),
        (A['rfr'],"  Risk-Free Rate (%)","rfr",FMTP),
        (A['erp'],"  Equity Risk Premium (%)","erp",FMTP),
        (A['beta'],"  Asset Beta","beta","#,##0.00"),
        (A['gearing'],"Gearing — Debt % of Capital","gearing",FMTP),
        (A['gamma'],"Gamma (AER Imputation Credits)","gamma","#,##0.000"),
        (A['tax_rate'],"Corporate Tax Rate (%)","tax_rate",FMTP)]:
        lbl(ws,row,lbl_t); inp(ws,row,DC,AV[key],fmt)
    lbl(ws,A['equity_pct'],"Equity % of Capital")
    fcell(ws,A['equity_pct'],DC,f"=1-B{A['gearing']}",FMTP)
    sec(ws,A['cpi_sec'],"CPI Inflation (%)")
    lbl(ws,A['cpi'],"CPI Rate (%)")
    for i in range(NH): hcell(ws,A['cpi'],DC+i,HV['cpi'][i],FMTP)
    for i in range(NF): inp(ws,A['cpi'],DC+NH+i,YV['cpi'][i],FMTP)
    sec(ws,A['vol_sec'],"Revenue & Volumes")
    H_GHW=[48000,49000,50000,51000,52000]; H_TAR=[12.0,12.3,12.7,13.0,13.2]
    for row,lbl_t,fkey,hdata,fmt in [
        (A['ghw'],"Transmission Volume (GWh)",'ghw',H_GHW,FMTN),
        (A['tariff'],"Average Tariff ($/MWh)",'tariff',H_TAR,"#,##0.00"),
        (A['conn_fees'],"Connection & Access Fees ($M)",'conn_fees',H_CONN,FMTD)]:
        lbl(ws,row,lbl_t)
        for i in range(NH): hcell(ws,row,DC+i,hdata[i],fmt)
        for i in range(NF): inp(ws,row,DC+NH+i,YV[fkey][i],fmt)
    sec(ws,A['opex_sec'],"Opex Forecast Allowances (Nominal $M)")
    for row,lbl_t,key in [(A['om'],"  Network O&M",'om'),(A['corp'],"  Corporate",'corp'),
        (A['insur'],"  Insurance",'insur'),(A['reg'],"  Regulatory",'reg')]:
        lbl(ws,row,lbl_t)
        for i in range(NF): inp(ws,row,DC+NH+i,YV[key][i],FMTD)
    sec(ws,A['capex_sec'],"Capex Programme — Forecast (Nominal $M)")
    for row,lbl_t,key in [(A['aug'],"  Augmentation",'aug'),(A['repl'],"  Replacement",'repl'),
        (A['conn'],"  Connections",'conn')]:
        lbl(ws,row,lbl_t)
        for i in range(NF): inp(ws,row,DC+NH+i,YV[key][i],FMTD)
    sec(ws,A['rab_sec'],"Regulatory Asset Base — Opening Balances")
    for row,lbl_t,key,fmt in [(A['open_rab'],"Opening RAB — FY2019 ($M)","open_rab",FMTD),
        (A['std_life'],"Standard Asset Life (years)","std_life",FMTN),
        (A['open_reg_dep'],"Opening Accum. Regulatory Dep. ($M)","open_reg_dep",FMTD)]:
        lbl(ws,row,lbl_t); inp(ws,row,DC,AV[key],fmt)
    sec(ws,A['ppe_sec'],"Accounting PP&E — Opening Balances (FY2019)")
    for row,lbl_t,key,fmt in [(A['open_ppe'],"Opening Gross PP&E ($M)","open_ppe",FMTD),
        (A['open_acc_da'],"Opening Accum. D&A ($M)","open_acc_da",FMTD),
        (A['acc_life'],"Accounting Asset Life (years)","acc_life",FMTN)]:
        lbl(ws,row,lbl_t); inp(ws,row,DC,AV[key],fmt)
    sec(ws,A['fin_sec'],"Financing Parameters")
    for row,lbl_t,key,fmt in [(A['open_debt'],"Opening LT Debt — FY2019 ($M)","open_debt",FMTD),
        (A['open_cash'],"Opening Cash — FY2019 ($M)","open_cash",FMTD),
        (A['cod'],"Nominal Cost of Debt (%)","cod",FMTP),
        (A['div_pay'],"Dividend Payout Ratio (%)","div_pay",FMTP)]:
        lbl(ws,row,lbl_t); inp(ws,row,DC,AV[key],fmt)
    for row,lbl_t,key in [(A['debt_iss'],"  Annual Debt Issuances ($M)",'debt_iss'),
        (A['debt_rep'],"  Annual Debt Repayments ($M)",'debt_rep')]:
        lbl(ws,row,lbl_t)
        for i in range(NF): inp(ws,row,DC+NH+i,YV[key][i],FMTD)
    sec(ws,A['wc_sec'],"Working Capital")
    for row,lbl_t,key in [(A['dso'],"Debtor Days (DSO)",'dso'),(A['dpo'],"Creditor Days (DPO)",'dpo')]:
        lbl(ws,row,lbl_t); inp(ws,row,DC,AV[key],FMTN)
    sec(ws,A['front_sec'],"Frontier Efficiency Parameters")
    for row,lbl_t,key in [(A['front_shift'],"Annual Frontier Shift (%)","front_shift"),
        (A['opex_eff'],"Opex Efficiency vs Frontier (%)","opex_eff"),
        (A['capex_eff'],"Capex Efficiency vs Frontier (%)","capex_eff")]:
        lbl(ws,row,lbl_t); inp(ws,row,DC,AV[key],FMTP)
    sec(ws,A['hist_opex_sec'],"Historical Opex — Actual ($M Nominal)")
    for row,lbl_t,key in [(A['h_om'],"  Network O&M",'om'),(A['h_corp'],"  Corporate",'corp'),
        (A['h_insur'],"  Insurance",'insur'),(A['h_reg'],"  Regulatory",'reg')]:
        lbl(ws,row,lbl_t)
        for i in range(NH): hcell(ws,row,DC+i,HV[key][i],FMTD)
    sec(ws,A['hist_capex_sec'],"Historical Capex — Actual ($M Nominal)")
    for row,lbl_t,key in [(A['h_aug'],"  Augmentation",'aug'),(A['h_repl'],"  Replacement",'repl'),
        (A['h_conn'],"  Connections",'conn')]:
        lbl(ws,row,lbl_t)
        for i in range(NH): hcell(ws,row,DC+i,HV[key][i],FMTD)
    sec(ws,A['hist_fin_sec'],"Historical Financing — Actual ($M)")
    for row,lbl_t,key in [(A['h_iss'],"  Debt Issuances",'debt_iss'),(A['h_rep'],"  Debt Repayments",'debt_rep')]:
        lbl(ws,row,lbl_t)
        for i in range(NH): hcell(ws,row,DC+i,HV[key][i],FMTD)
    sec(ws,A['hist_bs_sec'],"Opening Balance Sheet (FY2019) — $M")
    for row,lbl_t,key in [(A['open_re'],"Opening Retained Earnings ($M)","open_re"),
        (A['open_cap'],"Opening Share Capital ($M)","open_cap")]:
        lbl(ws,row,lbl_t); inp(ws,row,DC,AV[key],FMTD)


# ── DTM ────────────────────────────────────────────
def build_dtm(ws):
    ws.sheet_view.showGridLines=False
    DTM_DC=2; DTM_N=len(DTM_YRS)
    ws.column_dimensions['A'].width=44
    for i in range(DTM_N): ws.column_dimensions[get_column_letter(DTM_DC+i)].width=9
    ws.merge_cells(f"A1:{get_column_letter(DTM_DC+DTM_N-1)}1")
    c=ws.cell(row=1,column=1,value="DEBT TERM MODEL (DTM) — AER 10-Year Trailing Average Cost of Debt")
    c.fill=fill("003366"); c.font=mkf(bold=True,color="FFFFFF",size=13); c.alignment=LA
    ws.row_dimensions[1].height=26
    # Instructions
    ws.merge_cells(f"A{DT['instr']}:{get_column_letter(DTM_DC+DTM_N-1)}{DT['instr']}")
    c=ws.cell(row=DT['instr'],column=1,
        value="AER methodology: WACD = 10-year trailing average of benchmark bond yields (CGS + DRP + issuance costs). "
              "Each year 1/10th of total debt is re-priced at prevailing market rates.")
    c.fill=fill("FFF9C4"); c.font=mkf(italic=True,color="595959",size=9); c.alignment=LA
    # Year headers
    ws.row_dimensions[DT['hdr']].height=20
    c=ws.cell(row=DT['hdr'],column=1,value="% Rate"); c.fill=fill("003366"); c.font=F_H; c.alignment=LA
    for i,yr in enumerate(DTM_YRS):
        col=DTM_DC+i; suf="A" if yr<=2023 else "E"
        c=ws.cell(row=DT['hdr'],column=col,value=f"FY{yr}{suf}")
        c.fill=fill("003366"); c.font=F_YHH if yr<=2023 else F_YHF; c.alignment=CA
    # RFR section
    ws.cell(row=DT['rfr_sec'],column=1,value="COMPONENT YIELDS BY VINTAGE YEAR").font=F_S
    for c in range(1,DTM_DC+DTM_N): ws.cell(row=DT['rfr_sec'],column=c).fill=C_SEC
    lbl(ws,DT['rfr'],"  Risk-Free Rate — 10yr CGS Yield (%)")
    lbl(ws,DT['drp'],"  Debt Risk Premium — BBB+ Spread (%)")
    ws.cell(row=DT['yield_sec'],column=1,value="BENCHMARK YIELD CALCULATION").font=F_S
    for c in range(1,DTM_DC+DTM_N): ws.cell(row=DT['yield_sec'],column=c).fill=C_SEC
    lbl(ws,DT['bench_yield'],"  Benchmark Yield (RFR + DRP, %)")
    lbl(ws,DT['iss_cost'],"  + Issuance Costs (%)")
    lbl(ws,DT['total_yield'],"  Total Benchmark Cost of Debt (%)")
    for i,yr in enumerate(DTM_YRS):
        col=DTM_DC+i; cl=get_column_letter(col); is_hist=(yr<=2023)
        rfr_v=DTM_RFR[yr]/100; drp_v=DTM_DRP[yr]/100
        if is_hist:
            hcell(ws,DT['rfr'],col,rfr_v,FMTP2)
            hcell(ws,DT['drp'],col,drp_v,FMTP2)
        else:
            inp(ws,DT['rfr'],col,rfr_v,FMTP2)
            inp(ws,DT['drp'],col,drp_v,FMTP2)
        fcell(ws,DT['bench_yield'],col,f"={cl}{DT['rfr']}+{cl}{DT['drp']}",FMTP2,hist=is_hist)
        fcell(ws,DT['iss_cost'],col,ISS_COST,FMTP2,hist=is_hist)
        fcell(ws,DT['total_yield'],col,f"={cl}{DT['bench_yield']}+{cl}{DT['iss_cost']}",FMTP2,bold=True,hist=is_hist)
    # Trailing average section
    ws.cell(row=DT['trail_sec'],column=1,value="10-YEAR TRAILING AVERAGE WACD").font=F_S
    for c in range(1,DTM_DC+DTM_N): ws.cell(row=DT['trail_sec'],column=c).fill=C_SEC
    lbl(ws,DT['trail_avg'],"  Trailing Avg Cost of Debt (WACD, %)")
    lbl(ws,DT['n_vintages'],"  No. Vintage Years in Average")
    for i,yr in enumerate(DTM_YRS):
        col=DTM_DC+i; is_hist=(yr<=2023)
        # Use up to 10 prior vintage years
        n=min(i+1,10); start_i=max(0,i-9)
        start_cl=get_column_letter(DTM_DC+start_i); end_cl=get_column_letter(DTM_DC+i)
        fcell(ws,DT['trail_avg'],col,f"=AVERAGE({start_cl}{DT['total_yield']}:{end_cl}{DT['total_yield']})",FMTP2,bold=(yr>=2023),hist=is_hist)
        fcell(ws,DT['n_vintages'],col,n,FMTN,hist=is_hist)
    # Model year feed section (FY2019-FY2033 = DTM cols 6-25, i.e. DTM_YRS index 5-24)
    ws.cell(row=DT['model_sec'],column=1,value="WACD FEED TO FINANCIAL MODEL (FY2019A-FY2033E)").font=F_S
    for c in range(1,DTM_DC+DTM_N): ws.cell(row=DT['model_sec'],column=c).fill=C_SEC
    lbl(ws,DT['model_wacd'],"  Benchmark WACD for Model Year (%)")
    for i,yr in enumerate(DTM_YRS):
        col=DTM_DC+i; is_hist=(yr<=2023)
        if yr>=2019:
            fcell(ws,DT['model_wacd'],col,
                f"={get_column_letter(DTM_DC+i)}{DT['trail_avg']}",FMTP2,bold=True,hist=is_hist)
    ws.freeze_panes="C3"


# ── REFM ───────────────────────────────────────────
def build_refm(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"REVENUE & EXPENDITURE FORECASTING MODEL (REFM)","7030A0")
    yr_hdr(ws,RF['yr_hdr'])
    sec(ws,RF['opex_sec'],"Operating Expenditure ($M)")
    for row,lbl_t,hk,fk in [(RF['om'],"  Network O&M",'h_om','om'),(RF['corp'],"  Corporate",'h_corp','corp'),
        (RF['insur'],"  Insurance",'h_insur','insur'),(RF['reg'],"  Regulatory",'h_reg','reg')]:
        lbl(ws,row,lbl_t)
        for i in range(NH): fcell(ws,row,DC+i,xref("Assumptions",ylet(HIST[i]),A[hk]),FMTD,hist=True)
        for i in range(NF): fcell(ws,row,DC+NH+i,xref("Assumptions",ylet(FCST[i]),A[fk]),FMTD)
    lbl(ws,RF['tot_opex'],"Total Opex")
    for i in range(NT):
        col=DC+i; cl=get_column_letter(col); h=i<NH
        fcell(ws,RF['tot_opex'],col,f"=SUM({cl}{RF['om']}:{cl}{RF['reg']})",FMTD,bold=True,tot=True,hist=h)
    sec(ws,RF['capex_sec'],"Capital Expenditure ($M)")
    for row,lbl_t,hk,fk in [(RF['aug'],"  Augmentation",'h_aug','aug'),(RF['repl'],"  Replacement",'h_repl','repl'),
        (RF['conn'],"  Connections",'h_conn','conn')]:
        lbl(ws,row,lbl_t)
        for i in range(NH): fcell(ws,row,DC+i,xref("Assumptions",ylet(HIST[i]),A[hk]),FMTD,hist=True)
        for i in range(NF): fcell(ws,row,DC+NH+i,xref("Assumptions",ylet(FCST[i]),A[fk]),FMTD)
    lbl(ws,RF['tot_capex'],"Total Capex")
    for i in range(NT):
        col=DC+i; cl=get_column_letter(col); h=i<NH
        fcell(ws,RF['tot_capex'],col,f"=SUM({cl}{RF['aug']}:{cl}{RF['conn']})",FMTD,bold=True,tot=True,hist=h)
    sec(ws,RF['var_sec'],"Opex Variance Analysis — Allowed vs Actual ($M)")
    lbl(ws,RF['opex_allowed'],"  Opex Allowed (from PTRM)")
    lbl(ws,RF['opex_actual'],"  Opex Actual (this sheet)")
    lbl(ws,RF['opex_var'],"  Variance (Allowed − Actual)")
    for i in range(NT):
        col=DC+i; cl=get_column_letter(col); h=i<NH
        fcell(ws,RF['opex_allowed'],col,xref("PTRM",cl,PT['opex_all']),FMTD,hist=h)
        fcell(ws,RF['opex_actual'],col,f"={cl}{RF['tot_opex']}",FMTD,hist=h)
        fcell(ws,RF['opex_var'],col,f"={cl}{RF['opex_allowed']}-{cl}{RF['opex_actual']}",FMTD,bold=True,hist=h)
    sec(ws,RF['ecm_sec'],"Efficiency Carryover Mechanism (ECM)")
    lbl(ws,RF['ecm_rate'],"  ECM Sharing Rate"); inp(ws,RF['ecm_rate'],DC,0.30,FMTP)
    lbl(ws,RF['ecm_ben'],"  ECM Benefit ($M)")
    for i in range(NT):
        col=DC+i; cl=get_column_letter(col); h=i<NH
        fcell(ws,RF['ecm_ben'],col,f"=MAX({cl}{RF['opex_var']},0)*$B${RF['ecm_rate']}",FMTD,hist=h)

# ── Frontier ───────────────────────────────────────
def build_frontier(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"FRONTIER PRICING MODEL — Efficiency Shift & X-Factor","FF6600")
    yr_hdr(ws,FR['yr_hdr'])
    sec(ws,FR['sec'],"Frontier Efficiency Adjustment")
    for lbl_t,key in [("  Cumulative Frontier Shift Factor",FR['cum_shift']),
        ("  Raw Opex — from REFM ($M)",FR['raw_opex']),
        ("  Frontier-Adjusted Opex Allowance ($M)",FR['adj_opex']),
        ("  Raw Capex — from REFM ($M)",FR['raw_capex']),
        ("  Frontier-Adjusted Capex ($M)",FR['adj_capex'])]:
        lbl(ws,key,lbl_t)
    lbl(ws,FR['xfactor'],"  X-Factor (%)")
    for i in range(NT):
        col=DC+i; cl=get_column_letter(col); h=i<NH; fi=i-NH+1
        if h:
            fcell(ws,FR['cum_shift'],col,1.0,FMTP2,hist=True)
            fcell(ws,FR['adj_opex'],col,xref("REFM",cl,RF['tot_opex']),FMTD,hist=True)
            fcell(ws,FR['adj_capex'],col,xref("REFM",cl,RF['tot_capex']),FMTD,hist=True)
            fcell(ws,FR['xfactor'],col,0.0,FMTP,hist=True)
        else:
            fcell(ws,FR['cum_shift'],col,f"=(1-Assumptions!$B${A['front_shift']})^{fi}",FMTP2)
            fcell(ws,FR['adj_opex'],col,
                f"=REFM!{cl}{RF['tot_opex']}*{cl}{FR['cum_shift']}*Assumptions!$B${A['opex_eff']}",FMTD)
            fcell(ws,FR['adj_capex'],col,
                f"=REFM!{cl}{RF['tot_capex']}*Assumptions!$B${A['capex_eff']}",FMTD)
            fcell(ws,FR['xfactor'],col,f"=Assumptions!$B${A['front_shift']}",FMTP)
        fcell(ws,FR['raw_opex'],col,xref("REFM",cl,RF['tot_opex']),FMTD,hist=h)
        fcell(ws,FR['raw_capex'],col,xref("REFM",cl,RF['tot_capex']),FMTD,hist=h)


# ── RAB ────────────────────────────────────────────
def build_rab(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"REGULATORY ASSET BASE (RAB) ROLL-FORWARD","7B3F00")
    yr_hdr(ws,RB['yr_hdr'])
    sec(ws,RB['sec'],"RAB Roll-Forward ($M)")
    for lbl_t,key in [("  Opening RAB",RB['open_rab']),("  + CPI Indexation",RB['cpi_idx']),
        ("  + Capex Additions",RB['capex']),("  − Regulatory Depreciation",RB['reg_dep']),
        ("  = Closing RAB",RB['close_rab'])]:
        lbl(ws,key,lbl_t)
    lbl(ws,RB['avg_rab'],"  Average RAB (Return on RAB base)")
    sec(ws,RB['accum_sec'],"Accumulated Regulatory Depreciation ($M)")
    lbl(ws,RB['accum_reg_dep'],"  Accum. Reg. Depreciation (closing)")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        pv=get_column_letter(col-1) if col>DC else None
        if yr==2019:
            fcell(ws,RB['open_rab'],col,f"=Assumptions!$B${A['open_rab']}",FMTD,hist=True)
        else:
            fcell(ws,RB['open_rab'],col,f"={pv}{RB['close_rab']}",FMTD,hist=h)
        fcell(ws,RB['cpi_idx'],col,f"={cl}{RB['open_rab']}*Assumptions!{cl}{A['cpi']}",FMTD,hist=h)
        fcell(ws,RB['capex'],col,xref("REFM",cl,RF['tot_capex']),FMTD,hist=h)
        fcell(ws,RB['reg_dep'],col,f"={cl}{RB['open_rab']}/Assumptions!$B${A['std_life']}",FMTD,hist=h)
        fcell(ws,RB['close_rab'],col,
            f"={cl}{RB['open_rab']}+{cl}{RB['cpi_idx']}+{cl}{RB['capex']}-{cl}{RB['reg_dep']}",
            FMTD,bold=True,tot=True,hist=h)
        fcell(ws,RB['avg_rab'],col,f"=({cl}{RB['open_rab']}+{cl}{RB['close_rab']})/2",FMTD,hist=h)
        if yr==2019:
            fcell(ws,RB['accum_reg_dep'],col,
                f"=Assumptions!$B${A['open_reg_dep']}+{cl}{RB['reg_dep']}",FMTD,hist=True)
        else:
            fcell(ws,RB['accum_reg_dep'],col,
                f"={pv}{RB['accum_reg_dep']}+{cl}{RB['reg_dep']}",FMTD,hist=h)

# ── PTRM ───────────────────────────────────────────
def build_ptrm(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"POST TAX REVENUE MODEL (PTRM) — Maximum Allowed Revenue","0070C0")
    yr_hdr(ws,PT['yr_hdr'])
    sec(ws,PT['sec'],"MAR Build-Up ($M)")
    for lbl_t,key in [("  Return on RAB  (Avg RAB × WACC)",PT['ret_on']),
        ("  Return of RAB  (Regulatory Depreciation)",PT['ret_of']),
        ("  Opex Allowance  (Frontier-adjusted)",PT['opex_all']),
        ("  Tax Allowance",PT['tax_all']),
        ("  Maximum Allowed Revenue (MAR)",PT['mar'])]:
        lbl(ws,key,lbl_t)
    sec(ws,PT['tax_sec'],"Tax Allowance Build-Up ($M)")
    for lbl_t,key in [("  Taxable Income (Return on RAB)",PT['taxable']),
        ("  Gross Tax (× tax rate)",PT['tax_gross']),
        ("  Gamma Benefit (× gamma × tax rate)",PT['gamma_ben']),
        ("  Net Tax Allowance",PT['net_tax'])]:
        lbl(ws,key,lbl_t)
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST; idx=yr-2019
        if h:
            d=HIST_DATA[idx]
            avg_rab=d['avg_rab']
            ret_on=round(avg_rab*AV['wacc'],1)
            ret_of=d['reg_dep']
            opex_all=d['tot_op']
            taxable=ret_on
            tax_g=round(taxable*AV['tax_rate'],1)
            gamma_b=round(taxable*AV['gamma']*AV['tax_rate'],1)
            net_t=round(tax_g-gamma_b,1)
            mar=d['rev']
            hcell(ws,PT['ret_on'],col,ret_on,FMTD)
            hcell(ws,PT['ret_of'],col,ret_of,FMTD)
            hcell(ws,PT['opex_all'],col,opex_all,FMTD)
            hcell(ws,PT['tax_all'],col,net_t,FMTD)
            hcell(ws,PT['mar'],col,mar,FMTD,bold=True)
            hcell(ws,PT['taxable'],col,taxable,FMTD)
            hcell(ws,PT['tax_gross'],col,tax_g,FMTD)
            hcell(ws,PT['gamma_ben'],col,gamma_b,FMTD)
            hcell(ws,PT['net_tax'],col,net_t,FMTD)
        else:
            fcell(ws,PT['taxable'],col,f"=RAB!{cl}{RB['avg_rab']}*Assumptions!$B${A['wacc']}",FMTD)
            fcell(ws,PT['tax_gross'],col,f"={cl}{PT['taxable']}*Assumptions!$B${A['tax_rate']}",FMTD)
            fcell(ws,PT['gamma_ben'],col,
                f"={cl}{PT['taxable']}*Assumptions!$B${A['gamma']}*Assumptions!$B${A['tax_rate']}",FMTD)
            fcell(ws,PT['net_tax'],col,f"={cl}{PT['tax_gross']}-{cl}{PT['gamma_ben']}",FMTD)
            fcell(ws,PT['ret_on'],col,f"=RAB!{cl}{RB['avg_rab']}*Assumptions!$B${A['wacc']}",FMTD)
            fcell(ws,PT['ret_of'],col,xref("RAB",cl,RB['reg_dep']),FMTD)
            fcell(ws,PT['opex_all'],col,xref("Frontier",cl,FR['adj_opex']),FMTD)
            fcell(ws,PT['tax_all'],col,f"={cl}{PT['net_tax']}",FMTD)
            fcell(ws,PT['mar'],col,
                f"={cl}{PT['ret_on']}+{cl}{PT['ret_of']}+{cl}{PT['opex_all']}+{cl}{PT['tax_all']}",
                FMTD,bold=True,tot=True)


# ── Income Statement ───────────────────────────────
def build_is(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"INCOME STATEMENT — Electricity Grid Transmission Utility")
    yr_hdr(ws,IS['yr_hdr'])
    sec(ws,IS['rev_sec'],"Revenue ($M)")
    for lbl_t,key in [("  Transmission Revenue",IS['trans_rev']),
        ("  Connection & Access Fees",IS['conn_rev']),("  Total Revenue",IS['tot_rev'])]:
        lbl(ws,key,lbl_t)
    sec(ws,IS['opex_sec'],"Operating Expenditure ($M)")
    for lbl_t,key in [("  Network O&M",IS['om']),("  Corporate",IS['corp']),
        ("  Insurance",IS['insur']),("  Regulatory",IS['reg']),("  Total Opex",IS['tot_opex'])]:
        lbl(ws,key,lbl_t)
    for lbl_t,key in [("EBITDA",IS['ebitda']),("  Depreciation & Amortisation",IS['da']),
        ("EBIT",IS['ebit']),("  Interest Expense",IS['int_exp']),("EBT",IS['ebt']),
        ("  Income Tax",IS['tax']),("Net Income",IS['net_inc']),
        ("  Dividends",IS['divs']),("  Retained Earnings",IS['retd'])]:
        lbl(ws,key,lbl_t)
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST; idx=yr-2019
        pv=get_column_letter(col-1) if col>DC else None
        if h:
            d=HIST_DATA[idx]
            hcell(ws,IS['trans_rev'],col,d['trans'],FMTD)
            hcell(ws,IS['conn_rev'],col,d['conn'],FMTD)
            hcell(ws,IS['tot_rev'],col,d['rev'],FMTD,bold=True)
            for r_is,key in [(IS['om'],'om'),(IS['corp'],'corp'),(IS['insur'],'insur'),(IS['reg'],'reg')]:
                hcell(ws,r_is,col,HV[key][idx],FMTD)
            hcell(ws,IS['tot_opex'],col,d['tot_op'],FMTD,bold=True)
            hcell(ws,IS['ebitda'],col,d['ebitda'],FMTD,bold=True)
            hcell(ws,IS['da'],col,d['da'],FMTD)
            hcell(ws,IS['ebit'],col,d['ebit'],FMTD,bold=True)
            hcell(ws,IS['int_exp'],col,d['int_e'],FMTD)
            hcell(ws,IS['ebt'],col,d['ebt'],FMTD,bold=True)
            hcell(ws,IS['tax'],col,d['tax'],FMTD)
            hcell(ws,IS['net_inc'],col,d['ni'],FMTD,bold=True)
            hcell(ws,IS['divs'],col,d['divs'],FMTD)
            hcell(ws,IS['retd'],col,d['retd'],FMTD)
        else:
            fcell(ws,IS['conn_rev'],col,xref("Assumptions",cl,A['conn_fees']),FMTD)
            fcell(ws,IS['trans_rev'],col,f"=PTRM!{cl}{PT['mar']}-{cl}{IS['conn_rev']}",FMTD)
            fcell(ws,IS['tot_rev'],col,f"=PTRM!{cl}{PT['mar']}",FMTD,bold=True,sub=True)
            for r_is,r_rf in [(IS['om'],RF['om']),(IS['corp'],RF['corp']),
                (IS['insur'],RF['insur']),(IS['reg'],RF['reg'])]:
                fcell(ws,r_is,col,xref("REFM",cl,r_rf),FMTD)
            fcell(ws,IS['tot_opex'],col,f"=SUM({cl}{IS['om']}:{cl}{IS['reg']})",FMTD,bold=True,sub=True)
            fcell(ws,IS['ebitda'],col,f"={cl}{IS['tot_rev']}-{cl}{IS['tot_opex']}",FMTD,bold=True,tot=True)
            fcell(ws,IS['da'],col,
                f"='Balance Sheet'!{pv}{BS['gross_ppe']}/Assumptions!$B${A['acc_life']}",FMTD)
            fcell(ws,IS['ebit'],col,f"={cl}{IS['ebitda']}-{cl}{IS['da']}",FMTD,bold=True,tot=True)
            fcell(ws,IS['int_exp'],col,
                f"='Balance Sheet'!{pv}{BS['lt_debt']}*Assumptions!$B${A['cod']}",FMTD)
            fcell(ws,IS['ebt'],col,f"={cl}{IS['ebit']}-{cl}{IS['int_exp']}",FMTD,bold=True,tot=True)
            fcell(ws,IS['tax'],col,f"=MAX({cl}{IS['ebt']}*Assumptions!$B${A['tax_rate']},0)",FMTD)
            fcell(ws,IS['net_inc'],col,f"={cl}{IS['ebt']}-{cl}{IS['tax']}",FMTD,bold=True,tot=True)
            fcell(ws,IS['divs'],col,f"={cl}{IS['net_inc']}*Assumptions!$B${A['div_pay']}",FMTD)
            fcell(ws,IS['retd'],col,f"={cl}{IS['net_inc']}-{cl}{IS['divs']}",FMTD)


# ── Balance Sheet ──────────────────────────────────
def build_bs(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"BALANCE SHEET — Electricity Grid Transmission Utility")
    yr_hdr(ws,BS['yr_hdr'])
    sec(ws,BS['asset_sec'],"ASSETS ($M)")
    for lbl_t,key in [("  Cash & Cash Equivalents",BS['cash']),("  Accounts Receivable",BS['ar']),
        ("  Total Current Assets",BS['tot_ca']),("  Gross PP&E",BS['gross_ppe']),
        ("  (Less) Accumulated D&A",BS['acc_da']),("  Net PP&E",BS['net_ppe']),
        ("  Regulatory Asset (RAB excess over book)",BS['reg_asset']),
        ("  Total Non-Current Assets",BS['tot_nca']),("TOTAL ASSETS",BS['tot_assets'])]:
        lbl(ws,key,lbl_t)
    sec(ws,BS['liab_sec'],"LIABILITIES ($M)")
    for lbl_t,key in [("  Accounts Payable",BS['ap']),("  Total Current Liabilities",BS['tot_cl']),
        ("  Long-Term Debt",BS['lt_debt']),("  Deferred Tax Liability",BS['def_tax']),
        ("  Total Non-Current Liabilities",BS['tot_ncl']),("TOTAL LIABILITIES",BS['tot_liab'])]:
        lbl(ws,key,lbl_t)
    sec(ws,BS['eq_sec'],"EQUITY ($M)")
    for lbl_t,key in [("  Share Capital & Reserves",BS['sh_cap']),("  Retained Earnings",BS['ret_earn']),
        ("  Total Equity",BS['tot_eq']),("TOTAL LIABILITIES & EQUITY",BS['tot_le'])]:
        lbl(ws,key,lbl_t)
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST; idx=yr-2019
        pv=get_column_letter(col-1) if col>DC else None
        if h:
            d=HIST_DATA[idx]
            hcell(ws,BS['cash'],col,d['cash'],FMTD)
            hcell(ws,BS['ar'],col,d['ar'],FMTD)
            hcell(ws,BS['tot_ca'],col,round(d['cash']+d['ar'],1),FMTD,bold=True)
            hcell(ws,BS['gross_ppe'],col,d['ppe'],FMTD)
            hcell(ws,BS['acc_da'],col,d['acc_da'],FMTD)
            hcell(ws,BS['net_ppe'],col,d['net_ppe'],FMTD)
            hcell(ws,BS['reg_asset'],col,d['reg_asset'],FMTD)
            hcell(ws,BS['tot_nca'],col,round(d['net_ppe']+d['reg_asset'],1),FMTD,bold=True)
            tot_a=round(d['cash']+d['ar']+d['net_ppe']+d['reg_asset'],1)
            hcell(ws,BS['tot_assets'],col,tot_a,FMTD,bold=True)
            hcell(ws,BS['ap'],col,d['ap'],FMTD)
            hcell(ws,BS['tot_cl'],col,d['ap'],FMTD,bold=True)
            hcell(ws,BS['lt_debt'],col,d['lt_debt'],FMTD)
            hcell(ws,BS['def_tax'],col,d['def_tax'],FMTD)
            hcell(ws,BS['tot_ncl'],col,round(d['lt_debt']+d['def_tax'],1),FMTD,bold=True)
            tot_l=round(d['ap']+d['lt_debt']+d['def_tax'],1)
            hcell(ws,BS['tot_liab'],col,tot_l,FMTD,bold=True)
            hcell(ws,BS['sh_cap'],col,AV['open_cap'],FMTD)
            hcell(ws,BS['ret_earn'],col,d['re'],FMTD)
            hcell(ws,BS['tot_eq'],col,round(AV['open_cap']+d['re'],1),FMTD,bold=True)
            tot_le=round(tot_l+AV['open_cap']+d['re'],1)
            hcell(ws,BS['tot_le'],col,tot_le,FMTD,bold=True)
        else:
            fcell(ws,BS['cash'],col,xref("Cash Flow",cl,CF['end_cash']),FMTD)
            fcell(ws,BS['ar'],col,
                f"='Income Statement'!{cl}{IS['tot_rev']}*Assumptions!$B${A['dso']}/365",FMTD)
            fcell(ws,BS['tot_ca'],col,f"={cl}{BS['cash']}+{cl}{BS['ar']}",FMTD,bold=True,sub=True)
            fcell(ws,BS['gross_ppe'],col,
                f"={pv}{BS['gross_ppe']}+REFM!{cl}{RF['tot_capex']}",FMTD)
            fcell(ws,BS['acc_da'],col,
                f"={pv}{BS['acc_da']}+'Income Statement'!{cl}{IS['da']}",FMTD)
            fcell(ws,BS['net_ppe'],col,f"={cl}{BS['gross_ppe']}-{cl}{BS['acc_da']}",FMTD)
            fcell(ws,BS['reg_asset'],col,f"=RAB!{cl}{RB['close_rab']}-{cl}{BS['net_ppe']}",FMTD)
            fcell(ws,BS['tot_nca'],col,f"={cl}{BS['net_ppe']}+{cl}{BS['reg_asset']}",FMTD,bold=True,sub=True)
            fcell(ws,BS['tot_assets'],col,f"={cl}{BS['tot_ca']}+{cl}{BS['tot_nca']}",FMTD,bold=True,tot=True)
            fcell(ws,BS['ap'],col,
                f"=REFM!{cl}{RF['tot_opex']}*Assumptions!$B${A['dpo']}/365",FMTD)
            fcell(ws,BS['tot_cl'],col,f"={cl}{BS['ap']}",FMTD,bold=True,sub=True)
            fcell(ws,BS['lt_debt'],col,
                f"={pv}{BS['lt_debt']}+Assumptions!{cl}{A['debt_iss']}-Assumptions!{cl}{A['debt_rep']}",FMTD)
            fcell(ws,BS['def_tax'],col,
                f"=MAX((RAB!{cl}{RB['accum_reg_dep']}-{cl}{BS['acc_da']})*Assumptions!$B${A['tax_rate']},0)",FMTD)
            fcell(ws,BS['tot_ncl'],col,f"={cl}{BS['lt_debt']}+{cl}{BS['def_tax']}",FMTD,bold=True,sub=True)
            fcell(ws,BS['tot_liab'],col,f"={cl}{BS['tot_cl']}+{cl}{BS['tot_ncl']}",FMTD,bold=True,tot=True)
            fcell(ws,BS['sh_cap'],col,f"={pv}{BS['sh_cap']}",FMTD)
            fcell(ws,BS['ret_earn'],col,
                f"={pv}{BS['ret_earn']}+'Income Statement'!{cl}{IS['retd']}",FMTD)
            fcell(ws,BS['tot_eq'],col,f"={cl}{BS['sh_cap']}+{cl}{BS['ret_earn']}",FMTD,bold=True,sub=True)
            fcell(ws,BS['tot_le'],col,f"={cl}{BS['tot_liab']}+{cl}{BS['tot_eq']}",FMTD,bold=True,tot=True)


# ── Cash Flow ──────────────────────────────────────
def build_cf(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"CASH FLOW STATEMENT — Indirect Method")
    yr_hdr(ws,CF['yr_hdr'])
    sec(ws,CF['ocf_sec'],"OPERATING ACTIVITIES ($M)")
    for lbl_t,key in [("  Net Income",CF['net_inc']),("  + Depreciation & Amortisation",CF['da_add']),
        ("  +/− Change in Accounts Receivable",CF['d_ar']),
        ("  +/− Change in Accounts Payable",CF['d_ap']),
        ("  +/− Change in Regulatory Assets",CF['d_reg']),
        ("  Net Cash from Operating Activities",CF['tot_ocf'])]:
        lbl(ws,key,lbl_t)
    sec(ws,CF['icf_sec'],"INVESTING ACTIVITIES ($M)")
    for lbl_t,key in [("  Capital Expenditure",CF['capex']),("  Net Cash from Investing Activities",CF['tot_icf'])]:
        lbl(ws,key,lbl_t)
    sec(ws,CF['fcf_sec'],"FINANCING ACTIVITIES ($M)")
    for lbl_t,key in [("  Debt Issuances",CF['debt_iss']),("  Debt Repayments",CF['debt_rep']),
        ("  Dividends Paid",CF['divs_paid']),("  Net Cash from Financing Activities",CF['tot_fcf'])]:
        lbl(ws,key,lbl_t)
    for lbl_t,key in [("Net Change in Cash",CF['net_cf']),
        ("Opening Cash Balance",CF['open_cash']),("Closing Cash Balance",CF['end_cash'])]:
        lbl(ws,key,lbl_t)
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST; idx=yr-2019
        pv=get_column_letter(col-1) if col>DC else None
        if h:
            d=HIST_DATA[idx]
            hcell(ws,CF['net_inc'],col,d['ni'],FMTD)
            hcell(ws,CF['da_add'],col,d['da'],FMTD)
            hcell(ws,CF['d_ar'],col,round(d['d_ar'],1),FMTD)
            hcell(ws,CF['d_ap'],col,round(d['d_ap'],1),FMTD)
            hcell(ws,CF['d_reg'],col,0,FMTD)
            hcell(ws,CF['tot_ocf'],col,d['ocf'],FMTD,bold=True)
            hcell(ws,CF['capex'],col,-d['cap'],FMTD)
            hcell(ws,CF['tot_icf'],col,-d['cap'],FMTD,bold=True)
            hcell(ws,CF['debt_iss'],col,d['iss'],FMTD)
            hcell(ws,CF['debt_rep'],col,-d['rep'],FMTD)
            hcell(ws,CF['divs_paid'],col,-d['divs'],FMTD)
            hcell(ws,CF['tot_fcf'],col,d['fcf'],FMTD,bold=True)
            hcell(ws,CF['net_cf'],col,d['net_cf'],FMTD,bold=True)
            hcell(ws,CF['open_cash'],col,d['open_cash'],FMTD)
            hcell(ws,CF['end_cash'],col,d['end_cash'],FMTD,bold=True)
        else:
            fcell(ws,CF['net_inc'],col,xref("Income Statement",cl,IS['net_inc']),FMTD)
            fcell(ws,CF['da_add'],col,xref("Income Statement",cl,IS['da']),FMTD)
            fcell(ws,CF['d_ar'],col,
                f"=-('Balance Sheet'!{cl}{BS['ar']}-'Balance Sheet'!{pv}{BS['ar']})",FMTD)
            fcell(ws,CF['d_ap'],col,
                f"='Balance Sheet'!{cl}{BS['ap']}-'Balance Sheet'!{pv}{BS['ap']}",FMTD)
            fcell(ws,CF['d_reg'],col,
                f"=-('Balance Sheet'!{cl}{BS['reg_asset']}-'Balance Sheet'!{pv}{BS['reg_asset']})",FMTD)
            fcell(ws,CF['tot_ocf'],col,
                f"=SUM({cl}{CF['net_inc']}:{cl}{CF['d_reg']})",FMTD,bold=True,tot=True)
            fcell(ws,CF['capex'],col,f"=-REFM!{cl}{RF['tot_capex']}",FMTD)
            fcell(ws,CF['tot_icf'],col,f"={cl}{CF['capex']}",FMTD,bold=True,tot=True)
            fcell(ws,CF['debt_iss'],col,f"=Assumptions!{cl}{A['debt_iss']}",FMTD)
            fcell(ws,CF['debt_rep'],col,f"=-Assumptions!{cl}{A['debt_rep']}",FMTD)
            fcell(ws,CF['divs_paid'],col,f"=-'Income Statement'!{cl}{IS['divs']}",FMTD)
            fcell(ws,CF['tot_fcf'],col,
                f"={cl}{CF['debt_iss']}+{cl}{CF['debt_rep']}+{cl}{CF['divs_paid']}",FMTD,bold=True,tot=True)
            fcell(ws,CF['net_cf'],col,
                f"={cl}{CF['tot_ocf']}+{cl}{CF['tot_icf']}+{cl}{CF['tot_fcf']}",FMTD,bold=True,tot=True)
            if i==NH:  # First forecast year: opening = last hist BS cash
                last_h=ylet(2023)
                fcell(ws,CF['open_cash'],col,f"='Balance Sheet'!{last_h}{BS['cash']}",FMTD)
            else:
                fcell(ws,CF['open_cash'],col,f"={pv}{CF['end_cash']}",FMTD)
            fcell(ws,CF['end_cash'],col,
                f"={cl}{CF['open_cash']}+{cl}{CF['net_cf']}",FMTD,bold=True,tot=True)

# ── Checks ─────────────────────────────────────────
def build_checks(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"MODEL CHECKS — Balance Sheet | CFS Roll-Forward | RAB Integrity","C00000")
    yr_hdr(ws,2)
    sec(ws,4,"BALANCE SHEET CHECK")
    for r,lbl_t,sh,sr in [(5,"  Total Assets","Balance Sheet",BS['tot_assets']),
        (6,"  Total Liabilities & Equity","Balance Sheet",BS['tot_le'])]:
        lbl(ws,r,lbl_t)
        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            fcell(ws,r,col,xref(sh,cl,sr),FMTD,hist=h)
    lbl(ws,7,"  Discrepancy (target = 0)")
    sec(ws,9,"CFS ROLL-FORWARD CHECK")
    for r,lbl_t,sh,sr in [(10,"  Closing Cash (Balance Sheet)","Balance Sheet",BS['cash']),
        (11,"  Closing Cash (Cash Flow)","Cash Flow",CF['end_cash'])]:
        lbl(ws,r,lbl_t)
        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            fcell(ws,r,col,xref(sh,cl,sr),FMTD,hist=h)
    lbl(ws,12,"  Discrepancy (target = 0)")
    sec(ws,14,"RAB INTEGRITY CHECK")
    lbl(ws,15,"  Closing RAB (RAB sheet)")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws,15,col,xref("RAB",cl,RB['close_rab']),FMTD,hist=h)
    lbl(ws,16,"  Discrepancy (target = 0)")
    # Check formulas
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        for r,f in [(7,f"={cl}5-{cl}6"),(12,f"={cl}10-{cl}11"),(16,f"={cl}15-{cl}15")]:
            c=ws.cell(row=r,column=col,value=f)
            c.number_format='#,##0;[RED](#,##0)'; c.alignment=RA
            c.fill=C_HIS if h else C_OK; c.font=F_B


# ── Main ───────────────────────────────────────────
def build():
    wb = Workbook()
    ws_cover = wb.active; ws_cover.title="Cover"
    ws_assm  = wb.create_sheet("Assumptions")
    ws_dtm   = wb.create_sheet("DTM")
    ws_refm  = wb.create_sheet("REFM")
    ws_front = wb.create_sheet("Frontier")
    ws_rab   = wb.create_sheet("RAB")
    ws_ptrm  = wb.create_sheet("PTRM")
    ws_is    = wb.create_sheet("Income Statement")
    ws_bs    = wb.create_sheet("Balance Sheet")
    ws_cf    = wb.create_sheet("Cash Flow")
    ws_chk   = wb.create_sheet("Checks")

    build_cover(ws_cover)
    build_assumptions(ws_assm)
    build_dtm(ws_dtm)
    build_refm(ws_refm)
    build_frontier(ws_front)
    build_rab(ws_rab)
    build_ptrm(ws_ptrm)
    build_is(ws_is)
    build_bs(ws_bs)
    build_cf(ws_cf)
    build_checks(ws_chk)

    # Tab colours
    ws_cover.sheet_properties.tabColor = "001F4D"
    ws_assm.sheet_properties.tabColor  = "1B6B1B"
    ws_dtm.sheet_properties.tabColor   = "0070C0"
    ws_refm.sheet_properties.tabColor  = "7030A0"
    ws_front.sheet_properties.tabColor = "FF6600"
    ws_rab.sheet_properties.tabColor   = "7B3F00"
    ws_ptrm.sheet_properties.tabColor  = "003366"
    ws_is.sheet_properties.tabColor    = "002060"
    ws_bs.sheet_properties.tabColor    = "002060"
    ws_cf.sheet_properties.tabColor    = "002060"
    ws_chk.sheet_properties.tabColor   = "C00000"

    wb.save("three_way_finance_model.xlsx")
    tabs=[ws.title for ws in wb.worksheets]
    print(f"✓  three_way_finance_model.xlsx created — {len(tabs)} tabs: {tabs}")

if __name__ == "__main__":
    build()
