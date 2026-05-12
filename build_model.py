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



# ══════════════════════════════════════════════════
# NEW TAB ROW MAPS
# ══════════════════════════════════════════════════
PR = {}  # Projects
CP = {}  # Capex
FA = {}  # Fixed Assets
WF = {}  # Workforce
MN = {}  # Maintenance
IT = {}  # Initiatives

PR.update({'title':1,'yr_hdr':2,'sec':4,'hdr_row':5,
    'p001':6,'p002':7,'p003':8,'p004':9,'p005':10,
    'p006':11,'p007':12,'p008':13,'p009':14,'p010':15,
    'tot_capex':17,'aug':19,'repl':20,'conn':21})

CP.update({'title':1,'yr_hdr':2,'sec':4,'aug':5,'repl':6,'conn':7,'tot_proj':8,
    'ovhd_sec':10,'ovhd_rate':11,'ovhd_amt':12,'tot_capex':13,
    'recon_sec':15,'assm_capex':16,'variance':17})

# Fixed Assets: 5 asset classes × 7 rows each + 2 blank = ~42 rows; totals below
_FA_CLASSES = ['lines','subs','trans','scada','other']
_FA_BASE = {'lines':5,'subs':13,'trans':21,'scada':29,'other':37}
for _cls, _base in _FA_BASE.items():
    FA[f'{_cls}_sec']   = _base
    FA[f'{_cls}_ogross'] = _base+1
    FA[f'{_cls}_adds']   = _base+2
    FA[f'{_cls}_disp']   = _base+3
    FA[f'{_cls}_cgross']  = _base+4
    FA[f'{_cls}_oda']    = _base+5
    FA[f'{_cls}_da_chg'] = _base+6
    FA[f'{_cls}_cda']    = _base+7
FA.update({'title':1,'yr_hdr':2,'asset_sec':4,'tot_sec':46,
    'tot_gross':47,'tot_accum_da':48,'tot_net_ppe':49,'da_charge':50})

# Workforce: 5 departments × 5 rows + totals
_WF_DEPTS = ['ops','eng','corp','ict','exec']
_WF_BASE = {'ops':8,'eng':15,'corp':22,'ict':29,'exec':36}
for _d, _b in _WF_BASE.items():
    WF[f'{_d}_sec']   = _b
    WF[f'{_d}_fte']   = _b+1
    WF[f'{_d}_sal']   = _b+2
    WF[f'{_d}_cost']  = _b+3
    WF[f'{_d}_oncost']= _b+4
    WF[f'{_d}_total'] = _b+5
WF.update({'title':1,'yr_hdr':2,'oncost_sec':4,'oncost_rate':5,
    'dept_sec':7,'tot_sec':44,'tot_fte':45,'tot_cost':46,
    'ops_feed':47,'supp_feed':48})

MN.update({'title':1,'yr_hdr':2,'plan_sec':4,'lines':5,'subs':6,'trans':7,
    'scada':8,'other':9,'tot_planned':10,
    'corr_sec':12,'corr_rate':13,'tot_corr':14,
    'tot_sec':16,'tot_maint':17})

IT.update({'title':1,'yr_hdr':2,
    'rev_sec':4,'rev_ancillary':5,'rev_conn':6,'rev_tariff':7,'tot_rev':8,
    'opex_sec':10,'opex_procure':11,'opex_auto':12,'opex_energy':13,'tot_opex':14,
    'fin_sec':16,'fin_refi':17,'fin_hedge':18,'tot_fin':19,
    'dep_sec':21,'dep_life':22,'dep_disp':23,'tot_dep':24,
    'net_sec':26,'net_value':27})


# ── Projects ───────────────────────────────────────
def build_projects(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws,46,11); frz(ws)
    title_row(ws,"CAPITAL PROJECTS PIPELINE — Project-Level Capex Schedule","1F5C8B")
    yr_hdr(ws, PR['yr_hdr'])
    sec(ws,PR['sec'],"Capital Projects ($M Nominal) — Forecast Period FY2024E–FY2033E")
    # Column headers for label section
    ws.cell(row=PR['hdr_row'],column=1,value="  Project Name  |  Type  |  Budget $M").font=F_B
    for col in range(DC,DC+NT): ws.cell(row=PR['hdr_row'],column=col).fill=C_SEC

    PROJECTS = [
        # (id_key, name, type_tag, budget, {yr:$M per yr})
        ('p001','P001  Grid Modernisation Program','AUG', 630,
          {2024:90,2025:130,2026:150,2027:120,2028:80,2029:60}),
        ('p002','P002  Substation Upgrades — Zone A','REPL',210,
          {2024:70,2025:90,2026:50}),
        ('p003','P003  HVDC Interconnect (Stage 1)','AUG', 480,
          {2025:60,2026:120,2027:150,2028:120,2029:30}),
        ('p004','P004  New Connection — Wind Farm A','CONN',45,
          {2024:25,2025:20}),
        ('p005','P005  Transformer Replacement Program','REPL',350,
          {2024:35,2025:40,2026:45,2027:50,2028:55,2029:55,2030:50,2031:20}),
        ('p006','P006  Protection Relay Upgrade','REPL',120,
          {2024:30,2025:35,2026:30,2027:25}),
        ('p007','P007  New Connection — Solar Farm B','CONN',80,
          {2025:40,2026:40}),
        ('p008','P008  Transmission Reconductoring','REPL',420,
          {2026:50,2027:70,2028:90,2029:90,2030:80,2031:40}),
        ('p009','P009  Digital Twin Infrastructure','AUG', 90,
          {2027:30,2028:35,2029:25}),
        ('p010','P010  Reactive Power Support','AUG', 75,
          {2028:25,2029:30,2030:20}),
    ]
    TYPE_COLOR={'AUG':'DCE6F1','REPL':'FFF2CC','CONN':'E2EFDA'}

    for proj_key,name,ptype,budget,sched in PROJECTS:
        row=PR[proj_key]
        lbl(ws,row,name)
        ws.cell(row=row,column=1).fill=fill(TYPE_COLOR.get(ptype,'FFFFFF'))
        for yr,amt in sched.items():
            if 2024<=yr<=2033:
                col=ycol(yr)
                fcell(ws,row,col,amt,FMTD)

    # Legend
    for col in range(DC,DC+NT): ws.cell(row=PR['hdr_row']-1,column=col).fill=C_HDR

    # Total capex by year
    lbl(ws,PR['tot_capex'],"TOTAL CAPEX BY YEAR ($M)")
    ws.cell(row=PR['tot_capex'],column=1).font=F_B
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        if h:
            val=HV['aug'][i]+HV['repl'][i]+HV['conn'][i]
            hcell(ws,PR['tot_capex'],col,val,FMTD,bold=True)
        else:
            row_refs=",".join(f"{cl}{PR[k]}" for k in
                ['p001','p002','p003','p004','p005','p006','p007','p008','p009','p010'])
            fcell(ws,PR['tot_capex'],col,f"=SUM({row_refs})",FMTD,bold=True,tot=True)

    # By type
    TYPE_TAGS={'aug':'AUG','repl':'REPL','conn':'CONN'}
    for row,lbl_t,ptype in [(PR['aug'],"  Augmentation (AUG)",'AUG'),
        (PR['repl'],"  Replacement (REPL)",'REPL'),
        (PR['conn'],"  Connections (CONN)",'CONN')]:
        lbl(ws,row,lbl_t)
        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            if h:
                type_key={'AUG':'aug','REPL':'repl','CONN':'conn'}[ptype]
                hcell(ws,row,col,HV[type_key][i],FMTD,bold=False)
            else:
                matching=[k for k,(_,t,_,_) in
                    [(p[0],p[1:]) for p in PROJECTS] if t==ptype
                    ] if False else []
                # Compute from projects
                proj_total = sum(
                    p[4].get(yr,0) for p in PROJECTS if p[2]==ptype)
                fcell(ws,row,col,proj_total if proj_total>0 else 0,FMTD)

    # Note
    note=ws.cell(row=PR['tot_capex']+1,column=1,
        value="  AUG=Augmentation (blue) | REPL=Replacement (yellow) | CONN=Connections (green)")
    note.font=mkf(italic=True,color="595959",size=9)


# ── Capex ──────────────────────────────────────────
def build_capex(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"CAPITAL EXPENDITURE SCHEDULE — Detailed Capex Build-Up","1F5C8B")
    yr_hdr(ws, CP['yr_hdr'])
    sec(ws,CP['sec'],"Project-Sourced Capex by Category ($M Nominal)")
    # Pull from Projects tab (aug/repl/conn rows there)
    for row,lbl_t,pr_row in [(CP['aug'],"  Augmentation Capex",PR['aug']),
        (CP['repl'],"  Replacement / Renewal",PR['repl']),
        (CP['conn'],"  Connections",PR['conn'])]:
        lbl(ws,row,lbl_t)
        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            fcell(ws,row,col,xref("Projects",cl,pr_row),FMTD,hist=h)
    lbl(ws,CP['tot_proj'],"  Total Project Capex")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws,CP['tot_proj'],col,
            f"={cl}{CP['aug']}+{cl}{CP['repl']}+{cl}{CP['conn']}",FMTD,bold=True,sub=True,hist=h)

    sec(ws,CP['ovhd_sec'],"Capitalised Overheads")
    lbl(ws,CP['ovhd_rate'],"  Overhead Capitalisation Rate (%)")
    inp(ws,CP['ovhd_rate'],DC,0.08,FMTP)
    lbl(ws,CP['ovhd_amt'],"  Capitalised Overheads ($M)")
    lbl(ws,CP['tot_capex'],"TOTAL CAPEX (incl. overheads)")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws,CP['ovhd_amt'],col,
            f"={cl}{CP['tot_proj']}*$B${CP['ovhd_rate']}",FMTD,hist=h)
        fcell(ws,CP['tot_capex'],col,
            f"={cl}{CP['tot_proj']}+{cl}{CP['ovhd_amt']}",FMTD,bold=True,tot=True,hist=h)

    sec(ws,CP['recon_sec'],"Reconciliation to Assumptions Capex ($M)")
    lbl(ws,CP['assm_capex'],"  Assumptions Total (aug+repl+conn)")
    lbl(ws,CP['variance'],"  Variance (Capex tab less Assumptions)")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        if h:
            assm_val=HV['aug'][i]+HV['repl'][i]+HV['conn'][i]
            hcell(ws,CP['assm_capex'],col,assm_val,FMTD)
            hcell(ws,CP['variance'],col,0,FMTD)
        else:
            assm=f"=Assumptions!{cl}{A['aug']}+Assumptions!{cl}{A['repl']}+Assumptions!{cl}{A['conn']}"
            fcell(ws,CP['assm_capex'],col,assm,FMTD)
            fcell(ws,CP['variance'],col,
                f"={cl}{CP['tot_capex']}-{cl}{CP['assm_capex']}",FMTD,bold=True)


# ── Fixed Assets ───────────────────────────────────
def build_fixed_assets(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws,46,11); frz(ws)
    title_row(ws,"FIXED ASSET REGISTER — PP&E Roll-Forward by Asset Class","4A4A8A")
    yr_hdr(ws, FA['yr_hdr'])
    sec(ws,FA['asset_sec'],"Gross PP&E and Accumulated D&A by Asset Class ($M)")

    CLASSES=[
        ('lines',  "TRANSMISSION LINES",   40, 3200, 1067, 0.50),
        ('subs',   "SUBSTATIONS",           40, 2400,  800, 0.30),
        ('trans',  "TRANSFORMERS",          30, 1200,  480, 0.10),
        ('scada',  "SCADA & CONTROL",       15,  600,  180, 0.07),
        ('other',  "OTHER EQUIPMENT",       20,  400,   73, 0.03),
    ]
    # (cls, label, life, opening_gross, opening_accum_da, capex_alloc_pct)

    for cls,label,life,og,oda,alloc in CLASSES:
        base=_FA_BASE[cls]
        sec_row=FA[f'{cls}_sec']
        ws.cell(row=sec_row,column=1,value=f"{label}  (Standard Life: {life} yrs)").font=F_S
        for c in range(1,DC+NT): ws.cell(row=sec_row,column=c).fill=C_SEC

        for sub_lbl,key in [("  Opening Gross PP&E",f'{cls}_ogross'),
            ("  + Additions (from Capex)",f'{cls}_adds'),
            ("  − Disposals",f'{cls}_disp'),
            ("  Closing Gross PP&E",f'{cls}_cgross'),
            ("  Opening Accum. D&A",f'{cls}_oda'),
            ("  D&A Charge (closing gross / life)",f'{cls}_da_chg'),
            ("  Closing Accum. D&A",f'{cls}_cda')]:
            lbl(ws,FA[key],sub_lbl)

        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            pv=get_column_letter(col-1) if col>DC else None

            if yr==2019:
                fcell(ws,FA[f'{cls}_ogross'],col,og,FMTD,hist=True)
                fcell(ws,FA[f'{cls}_oda'],   col,oda,FMTD,hist=True)
            else:
                fcell(ws,FA[f'{cls}_ogross'],col,f"={pv}{FA[f'{cls}_cgross']}",FMTD,hist=h)
                fcell(ws,FA[f'{cls}_oda'],   col,f"={pv}{FA[f'{cls}_cda']}",FMTD,hist=h)

            # Additions from Capex tab (alloc%)
            capex_src=f"='Capex'!{cl}{CP['tot_capex']}*{alloc}"
            fcell(ws,FA[f'{cls}_adds'],col,capex_src,FMTD,hist=h)
            fcell(ws,FA[f'{cls}_disp'],col,0,FMTD,hist=h)
            fcell(ws,FA[f'{cls}_cgross'],col,
                f"={cl}{FA[f'{cls}_ogross']}+{cl}{FA[f'{cls}_adds']}-{cl}{FA[f'{cls}_disp']}",
                FMTD,hist=h)
            # D&A = closing gross / life (simplified straight-line)
            fcell(ws,FA[f'{cls}_da_chg'],col,
                f"={cl}{FA[f'{cls}_cgross']}/{life}",FMTD,hist=h)
            fcell(ws,FA[f'{cls}_cda'],col,
                f"={cl}{FA[f'{cls}_oda']}+{cl}{FA[f'{cls}_da_chg']}",FMTD,hist=h)

    # Totals section
    sec(ws,FA['tot_sec'],"CONSOLIDATED TOTALS ($M)")
    for tot_lbl,tot_key,sub_key in [
        ("  Total Gross PP&E",FA['tot_gross'],f'{{}}_cgross'),
        ("  Total Accum. D&A",FA['tot_accum_da'],f'{{}}_cda'),
        ("  Total Net PP&E",  FA['tot_net_ppe'], None),
        ("  Total D&A Charge (feeds IS D&A)",FA['da_charge'],f'{{}}_da_chg')]:
        lbl(ws,tot_key,tot_lbl)
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        for tot_key,sk in [(FA['tot_gross'],'_cgross'),(FA['tot_accum_da'],'_cda'),(FA['da_charge'],'_da_chg')]:
            refs="+".join(f"{cl}{FA[f'{c}{sk}']}" for c in _FA_CLASSES)
            fcell(ws,tot_key,col,f"={refs}",FMTD,bold=True,hist=h)
        fcell(ws,FA['tot_net_ppe'],col,
            f"={cl}{FA['tot_gross']}-{cl}{FA['tot_accum_da']}",FMTD,bold=True,tot=True,hist=h)


# ── Workforce ──────────────────────────────────────
def build_workforce(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws,46,11); frz(ws)
    title_row(ws,"WORKFORCE — Headcount & Labour Cost Model","2E6B3E")
    yr_hdr(ws, WF['yr_hdr'])
    sec(ws,WF['oncost_sec'],"On-Cost Parameters")
    lbl(ws,WF['oncost_rate'],"  On-Cost Rate (super, leave, workers comp, etc.)")
    inp(ws,WF['oncost_rate'],DC,0.30,FMTP)

    DEPTS=[
        # (key, label, hist_fte[5], fcst_fte[10], avg_sal_base $000)
        ('ops',  "NETWORK OPERATIONS (field technicians & operators)",
          [248,251,254,257,260],[263,267,270,273,277,280,283,287,290,294],95),
        ('eng',  "ENGINEERING & PLANNING (asset, protection, planning)",
          [78,79,80,81,82],[83,84,86,87,89,90,92,93,95,97],120),
        ('corp', "CORPORATE & SHARED SERVICES (finance, HR, legal, comms)",
          [118,119,120,121,122],[123,124,125,126,127,128,129,130,131,132],110),
        ('ict',  "ICT & DIGITAL (infrastructure, cybersecurity, data)",
          [42,43,44,45,46],[48,50,52,54,55,57,58,59,60,61],115),
        ('exec', "EXECUTIVE & MANAGEMENT (C-suite, directors, managers)",
          [25,25,25,25,25],[25,25,25,25,25,25,25,26,26,26],250),
    ]

    sec(ws,WF['dept_sec'],"Workforce by Department")
    for dept_key,dept_lbl,h_fte,f_fte,sal_base in DEPTS:
        sec_row=WF[f'{dept_key}_sec']
        ws.cell(row=sec_row,column=1,value=dept_lbl).font=F_S
        for c in range(1,DC+NT): ws.cell(row=sec_row,column=c).fill=C_SEC
        lbl(ws,WF[f'{dept_key}_fte'],  "  Headcount (FTE)")
        lbl(ws,WF[f'{dept_key}_sal'],  "  Avg. Salary ($/FTE, nominal $'000)")
        lbl(ws,WF[f'{dept_key}_cost'], "  Salary Cost ($M)")
        lbl(ws,WF[f'{dept_key}_oncost'],"  On-Costs ($M)")
        lbl(ws,WF[f'{dept_key}_total'],"  Total Labour Cost ($M)")
        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            pv=get_column_letter(col-1) if col>DC else None
            fte=h_fte[i] if h else f_fte[i-NH]
            if h:
                hcell(ws,WF[f'{dept_key}_fte'],col,fte,FMTN)
                sal_nom=round(sal_base*(1+HV['cpi'][i])**i/1000,3)
                hcell(ws,WF[f'{dept_key}_sal'],col,round(sal_base*(1+0.025)**i,1),"#,##0")
            else:
                fcell(ws,WF[f'{dept_key}_fte'],col,fte,FMTN)
                fcell(ws,WF[f'{dept_key}_sal'],col,
                    f"={pv}{WF[f'{dept_key}_sal']}*(1+Assumptions!{cl}{A['cpi']})" if i>NH
                    else round(sal_base*(1+0.025)**NH,1),"#,##0",hist=False)
            fcell(ws,WF[f'{dept_key}_cost'],col,
                f"={cl}{WF[f'{dept_key}_fte']}*{cl}{WF[f'{dept_key}_sal']}/1000",FMTD,hist=h)
            fcell(ws,WF[f'{dept_key}_oncost'],col,
                f"={cl}{WF[f'{dept_key}_cost']}*$B${WF['oncost_rate']}",FMTD,hist=h)
            fcell(ws,WF[f'{dept_key}_total'],col,
                f"={cl}{WF[f'{dept_key}_cost']}+{cl}{WF[f'{dept_key}_oncost']}",
                FMTD,bold=True,hist=h)

    sec(ws,WF['tot_sec'],"TOTAL WORKFORCE SUMMARY")
    lbl(ws,WF['tot_fte'],"  Total Headcount (FTE)")
    lbl(ws,WF['tot_cost'],"  Total Labour Cost ($M)")
    lbl(ws,WF['ops_feed'],"  → Network Ops Cost (feeds O&M line)")
    lbl(ws,WF['supp_feed'],"  → Support Functions Cost (feeds Corporate line)")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fte_refs="+".join(f"{cl}{WF[f'{d}_fte']}" for d in _WF_DEPTS)
        cost_refs="+".join(f"{cl}{WF[f'{d}_total']}" for d in _WF_DEPTS)
        fcell(ws,WF['tot_fte'],col,f"={fte_refs}",FMTN,bold=True,hist=h)
        fcell(ws,WF['tot_cost'],col,f"={cost_refs}",FMTD,bold=True,tot=True,hist=h)
        fcell(ws,WF['ops_feed'],col,f"={cl}{WF['ops_total']}",FMTD,hist=h)
        supp_depts = ['eng','corp','ict','exec']
        supp_refs = '+'.join(f'{cl}{WF[d+"_total"]}' for d in supp_depts)
        fcell(ws,WF['supp_feed'],col,f'={supp_refs}',FMTD,hist=h)


# ── Maintenance ────────────────────────────────────
def build_maintenance(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws,46,11); frz(ws)
    title_row(ws,"MAINTENANCE — Asset Maintenance Cost Schedule","8B4513")
    yr_hdr(ws, MN['yr_hdr'])
    sec(ws,MN['plan_sec'],"PLANNED / PREVENTIVE MAINTENANCE ($M)")

    # Base maintenance rates per asset class ($M/yr, escalate with CPI)
    MAINT_BASE=[
        (MN['lines'], "  Transmission Lines",  24.0),
        (MN['subs'],  "  Substations",          18.0),
        (MN['trans'], "  Transformers",          12.0),
        (MN['scada'], "  SCADA & Control",        8.0),
        (MN['other'], "  Other Equipment",        4.5),
    ]
    H_CPI_CUM=[1.0,1.017,1.035,1.065,1.128]  # cumulative CPI index for FY2019-FY2023
    for row,lbl_t,base in MAINT_BASE:
        lbl(ws,row,lbl_t)
        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            pv=get_column_letter(col-1) if col>DC else None
            if h:
                hcell(ws,row,col,round(base*H_CPI_CUM[i],1),FMTD)
            else:
                if i==NH:
                    prev_val=round(base*H_CPI_CUM[-1],1)
                    fcell(ws,row,col,f"={prev_val}*(1+Assumptions!{cl}{A['cpi']})",FMTD)
                else:
                    fcell(ws,row,col,
                        f"={pv}{row}*(1+Assumptions!{cl}{A['cpi']})",FMTD)

    lbl(ws,MN['tot_planned'],"  Total Planned Maintenance")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws,MN['tot_planned'],col,
            f"={cl}{MN['lines']}+{cl}{MN['subs']}+{cl}{MN['trans']}+{cl}{MN['scada']}+{cl}{MN['other']}",
            FMTD,bold=True,sub=True,hist=h)

    sec(ws,MN['corr_sec'],"CORRECTIVE / EMERGENCY MAINTENANCE ($M)")
    lbl(ws,MN['corr_rate'],"  Corrective Rate (% of planned)")
    inp(ws,MN['corr_rate'],DC,0.15,FMTP)
    lbl(ws,MN['tot_corr'],"  Total Corrective Maintenance")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws,MN['tot_corr'],col,
            f"={cl}{MN['tot_planned']}*$B${MN['corr_rate']}",FMTD,hist=h)

    sec(ws,MN['tot_sec'],"TOTAL MAINTENANCE COST ($M) — Feeds Network O&M in IS")
    lbl(ws,MN['tot_maint'],"  Total Maintenance (Planned + Corrective)")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws,MN['tot_maint'],col,
            f"={cl}{MN['tot_planned']}+{cl}{MN['tot_corr']}",
            FMTD,bold=True,tot=True,hist=h)

    note=ws.cell(row=MN['tot_maint']+1,column=1,
        value="  Note: Maintenance total is a supporting schedule. "
              "IS Network O&M is currently driven by Assumptions/REFM. "
              "Update REFM O&M inputs to align with this schedule.")
    note.font=mkf(italic=True,color="595959",size=9)


# ── Initiatives ────────────────────────────────────
def build_initiatives(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws,48,11); frz(ws)
    title_row(ws,"STRATEGIC INITIATIVES — Revenue Growth & Cost Reduction Pipeline","B22222")
    yr_hdr(ws, IT['yr_hdr'])

    INIT_DATA={
        # Revenue: (row_key, label, {yr: annual_benefit $M})
        'rev':[
            ('rev_ancillary',"  Ancillary Services & FCAS Revenue",
              {2025:3,2026:5,2027:6,2028:7,2029:8,2030:8,2031:8,2032:9,2033:9}),
            ('rev_conn',    "  New Connection Pipeline (beyond P004/P007)",
              {2026:2,2027:3,2028:4,2029:4,2030:5,2031:5,2032:5,2033:6}),
            ('rev_tariff',  "  Tariff Structure & Access Review",
              {2027:4,2028:5,2029:5,2030:5,2031:6,2032:6,2033:6}),
        ],
        'opex':[
            ('opex_procure', "  Procurement & Supply Chain Savings",
              {2025:2,2026:3,2027:4,2028:5,2029:6,2030:6,2031:7,2032:7,2033:7}),
            ('opex_auto',    "  Digital Automation & Predictive Maintenance",
              {2026:1,2027:2,2028:3,2029:3,2030:4,2031:4,2032:5,2033:5}),
            ('opex_energy',  "  Energy Efficiency & Fleet Electrification",
              {2025:1,2026:1,2027:2,2028:2,2029:2,2030:2,2031:2,2032:2,2033:2}),
        ],
        'fin':[
            ('fin_refi',  "  Debt Refinancing — Lower Benchmark Rate",
              {2025:4,2026:5,2027:6,2028:7,2029:8,2030:8,2031:9,2032:9,2033:10}),
            ('fin_hedge', "  Interest Rate Hedging Programme",
              {2026:2,2027:2,2028:3,2029:3,2030:3,2031:3,2032:3,2033:3}),
        ],
        'dep':[
            ('dep_life',  "  Asset Life Extension Review (regulatory approved)",
              {2025:3,2026:4,2027:5,2028:5,2029:5,2030:5,2031:6,2032:6,2033:6}),
            ('dep_disp',  "  Targeted Asset Disposal & Write-Off",
              {2026:2,2027:2,2028:2,2029:2,2030:2,2031:2,2032:2,2033:2}),
        ],
    }
    CAT_META=[
        ('rev', IT['rev_sec'],  "REVENUE ENHANCEMENT ($M — annual benefit)",  IT['tot_rev'],  "C_HDR","→ Added to IS Revenue (above base MAR)"),
        ('opex',IT['opex_sec'], "OPEX REDUCTION INITIATIVES ($M — annual saving)",IT['tot_opex'],"7030A0","→ Reduces IS Total Opex"),
        ('fin', IT['fin_sec'],  "FINANCE COST REDUCTION ($M — annual saving)", IT['tot_fin'],  "FF6600","→ Reduces IS Interest Expense"),
        ('dep', IT['dep_sec'],  "DEPRECIATION REDUCTION ($M — annual saving)", IT['tot_dep'],  "7B3F00","→ Reduces IS D&A Charge"),
    ]

    for cat,sec_row,sec_lbl,tot_row,col_hex,feed_note in CAT_META:
        ws.cell(row=sec_row,column=1,value=sec_lbl).font=F_S
        for c in range(1,DC+NT): ws.cell(row=sec_row,column=c).fill=fill("F0F0F0")
        for init_key,init_lbl,sched in INIT_DATA[cat]:
            row=IT[init_key]; lbl(ws,row,init_lbl)
            for i,yr in enumerate(ALL):
                col=DC+i; h=yr in HIST
                val=sched.get(yr,0)
                if h:
                    hcell(ws,row,col,val if val else 0,FMTD)
                else:
                    fcell(ws,row,col,val if val else 0,FMTD)

        lbl(ws,tot_row,f"  TOTAL {cat.upper()} INITIATIVES")
        row_keys=[IT[k] for k,_,_ in INIT_DATA[cat]]
        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            refs="+".join(f"{cl}{r}" for r in row_keys)
            fcell(ws,tot_row,col,f"={refs}",FMTD,bold=True,tot=True,hist=h)
        # Feed note
        note=ws.cell(row=tot_row+1,column=1,value=f"  {feed_note}")
        note.font=mkf(italic=True,color="595959",size=9)

    # Net value section
    sec(ws,IT['net_sec'],"NET INITIATIVE VALUE — ANNUAL IMPACT ON NET INCOME ($M)")
    lbl(ws,IT['net_value'],"  Net Annual Value (Rev + Opex + Finance + D&A savings, pre-tax)")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws,IT['net_value'],col,
            f"={cl}{IT['tot_rev']}+{cl}{IT['tot_opex']}+{cl}{IT['tot_fin']}+{cl}{IT['tot_dep']}",
            FMTD,bold=True,tot=True,hist=h)


# ── Income Statement (updated — integrates Initiatives & Fixed Assets) ──
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
            # Revenue: PTRM MAR + revenue initiatives
            fcell(ws,IS['conn_rev'],col,xref("Assumptions",cl,A['conn_fees']),FMTD)
            fcell(ws,IS['trans_rev'],col,
                f"=PTRM!{cl}{PT['mar']}-{cl}{IS['conn_rev']}+Initiatives!{cl}{IT['tot_rev']}",FMTD)
            fcell(ws,IS['tot_rev'],col,
                f"=PTRM!{cl}{PT['mar']}+Initiatives!{cl}{IT['tot_rev']}",FMTD,bold=True,sub=True)
            # Opex: REFM minus opex initiatives
            for r_is,r_rf in [(IS['om'],RF['om']),(IS['corp'],RF['corp']),
                (IS['insur'],RF['insur']),(IS['reg'],RF['reg'])]:
                fcell(ws,r_is,col,xref("REFM",cl,r_rf),FMTD)
            fcell(ws,IS['tot_opex'],col,
                f"=SUM({cl}{IS['om']}:{cl}{IS['reg']})-Initiatives!{cl}{IT['tot_opex']}",
                FMTD,bold=True,sub=True)
            fcell(ws,IS['ebitda'],col,
                f"={cl}{IS['tot_rev']}-{cl}{IS['tot_opex']}",FMTD,bold=True,tot=True)
            # D&A from Fixed Assets tab minus depreciation reduction initiatives
            fcell(ws,IS['da'],col,
                f"='Fixed Assets'!{cl}{FA['da_charge']}-Initiatives!{cl}{IT['tot_dep']}",FMTD)
            fcell(ws,IS['ebit'],col,
                f"={cl}{IS['ebitda']}-{cl}{IS['da']}",FMTD,bold=True,tot=True)
            # Interest: prior debt × CoD minus finance initiatives
            fcell(ws,IS['int_exp'],col,
                f"='Balance Sheet'!{pv}{BS['lt_debt']}*Assumptions!$B${A['cod']}"
                f"-Initiatives!{cl}{IT['tot_fin']}",FMTD)
            fcell(ws,IS['ebt'],col,
                f"={cl}{IS['ebit']}-{cl}{IS['int_exp']}",FMTD,bold=True,tot=True)
            fcell(ws,IS['tax'],col,
                f"=MAX({cl}{IS['ebt']}*Assumptions!$B${A['tax_rate']},0)",FMTD)
            fcell(ws,IS['net_inc'],col,
                f"={cl}{IS['ebt']}-{cl}{IS['tax']}",FMTD,bold=True,tot=True)
            fcell(ws,IS['divs'],col,
                f"={cl}{IS['net_inc']}*Assumptions!$B${A['div_pay']}",FMTD)
            fcell(ws,IS['retd'],col,
                f"={cl}{IS['net_inc']}-{cl}{IS['divs']}",FMTD)

# ── REFM (updated — references Capex tab for forecast) ──
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
    sec(ws,RF['capex_sec'],"Capital Expenditure ($M) — Forecast sourced from Capex tab")
    for row,lbl_t,hk,cp_row in [
        (RF['aug'],"  Augmentation Capex",'h_aug',CP['aug']),
        (RF['repl'],"  Replacement / Renewal",'h_repl',CP['repl']),
        (RF['conn'],"  Connections",'h_conn',CP['conn'])]:
        lbl(ws,row,lbl_t)
        for i in range(NH): fcell(ws,row,DC+i,xref("Assumptions",ylet(HIST[i]),A[hk]),FMTD,hist=True)
        for i in range(NF): fcell(ws,row,DC+NH+i,xref("Capex",ylet(FCST[i]),cp_row),FMTD)
    lbl(ws,RF['tot_capex'],"Total Capex")
    for i in range(NT):
        col=DC+i; cl=get_column_letter(col); h=i<NH
        fcell(ws,RF['tot_capex'],col,f"=SUM({cl}{RF['aug']}:{cl}{RF['conn']})",FMTD,bold=True,tot=True,hist=h)
    sec(ws,RF['var_sec'],"Opex Variance — Allowed vs Actual ($M)")
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


# ── Main build function (17 tabs) ─────────────────────
def build():
    wb = Workbook()
    ws_cover = wb.active;          ws_cover.title = "Cover"
    ws_assm  = wb.create_sheet("Assumptions")
    ws_dtm   = wb.create_sheet("DTM")
    ws_proj  = wb.create_sheet("Projects")
    ws_capex = wb.create_sheet("Capex")
    ws_fa    = wb.create_sheet("Fixed Assets")
    ws_wf    = wb.create_sheet("Workforce")
    ws_mn    = wb.create_sheet("Maintenance")
    ws_refm  = wb.create_sheet("REFM")
    ws_front = wb.create_sheet("Frontier")
    ws_rab   = wb.create_sheet("RAB")
    ws_ptrm  = wb.create_sheet("PTRM")
    ws_init  = wb.create_sheet("Initiatives")
    ws_is    = wb.create_sheet("Income Statement")
    ws_bs    = wb.create_sheet("Balance Sheet")
    ws_cf    = wb.create_sheet("Cash Flow")
    ws_chk   = wb.create_sheet("Checks")

    build_cover(ws_cover)
    build_assumptions(ws_assm)
    build_dtm(ws_dtm)
    build_projects(ws_proj)
    build_capex(ws_capex)
    build_fixed_assets(ws_fa)
    build_workforce(ws_wf)
    build_maintenance(ws_mn)
    build_refm(ws_refm)
    build_frontier(ws_front)
    build_rab(ws_rab)
    build_ptrm(ws_ptrm)
    build_initiatives(ws_init)
    build_is(ws_is)
    build_bs(ws_bs)
    build_cf(ws_cf)
    build_checks(ws_chk)

    # Tab colours
    tab_colours = {
        "Cover":            "1F4E79",
        "Assumptions":      "375623",
        "DTM":              "7030A0",
        "Projects":         "BF8F00",
        "Capex":            "BF8F00",
        "Fixed Assets":     "BF8F00",
        "Workforce":        "BF8F00",
        "Maintenance":      "BF8F00",
        "REFM":             "7030A0",
        "Frontier":         "7030A0",
        "RAB":              "7030A0",
        "PTRM":             "7030A0",
        "Initiatives":      "C00000",
        "Income Statement": "C00000",
        "Balance Sheet":    "C00000",
        "Cash Flow":        "C00000",
        "Checks":           "808080",
    }
    for ws in wb.worksheets:
        colour = tab_colours.get(ws.title)
        if colour:
            ws.sheet_properties.tabColor = colour

    path = "three_way_finance_model.xlsx"
    wb.save(path)
    print(f"Saved: {path}  ({len(wb.worksheets)} sheets)")


# ════════════════════════════════════════════════════════════════
# RAB ROLL-FORWARD (DETAILED) + AER DEPRECIATION TRACKING MODELS
# ════════════════════════════════════════════════════════════════

# ── Row maps ─────────────────────────────────────────────────────
RM = {}   # RAB Roll-Forward (detailed)
DK = {}   # AER Depreciation Tracking

# RAB Roll-Forward: asset class data
# (key, label, AER_std_life_yrs, opening_rab_$M, capex_alloc_pct)
RM_CLASSES = [
    ('lines','TRANSMISSION LINES  |  AER Standard Life: 60 years',  60, 2400, 0.50),
    ('subs', 'SUBSTATIONS  |  AER Standard Life: 55 years',          55, 1440, 0.30),
    ('trans','TRANSFORMERS  |  AER Standard Life: 45 years',         45,  480, 0.10),
    ('scada','SCADA & CONTROL  |  AER Standard Life: 25 years',      25,  336, 0.07),
    ('other','OTHER EQUIPMENT  |  AER Standard Life: 30 years',      30,  144, 0.03),
]

# AER Depreciation Tracking: class data
# (key, AER_life, ATO_life, acc_life, opening_rab, opening_tab, alloc)
DK_CLASSES = [
    ('lines', 60, 40, 40, 2400, 1750, 0.50),
    ('subs',  55, 40, 40, 1440, 1050, 0.30),
    ('trans', 45, 30, 30,  480,  350, 0.10),
    ('scada', 25, 20, 15,  336,  245, 0.07),
    ('other', 30, 20, 20,  144,  105, 0.03),
]
BLENDED_ATO_LIFE = 37   # weighted avg: 0.5×40+0.3×40+0.1×30+0.07×20+0.03×20
_OPEN_ACCUM_TAX_DEP = 4300  # ≈ opening gross cost (7800) − opening TAB (3500)
_OPEN_DTL = 510             # (opening Net PPE 5200 − opening TAB 3500) × 30%

# RM row map: 5 classes × 6 rows each; bases at 4,11,18,25,32
_RM_BASES = {'lines':4,'subs':11,'trans':18,'scada':25,'other':32}
for _c, _b in _RM_BASES.items():
    RM[f'{_c}_sec']  = _b;   RM[f'{_c}_open'] = _b+1; RM[f'{_c}_cpi']  = _b+2
    RM[f'{_c}_cap']  = _b+3; RM[f'{_c}_dep']  = _b+4; RM[f'{_c}_close']= _b+5
RM.update({'title':1,'yr_hdr':2,
    'tot_sec':39,'tot_open':40,'tot_cpi':41,'tot_cap':42,'tot_dep':43,'tot_close':44,
    'avg_rab':46,
    'accum_sec':48,'accum_open':49,'accum_dep':50,'accum_close':51,
    'rcp_sec':53,'rcp1_cap':54,'rcp1_dep':55,'rcp1_avg':56,
    'rcp2_cap':58,'rcp2_dep':59,'rcp2_avg':60})

# DK row map
DK.update({'title':1,'yr_hdr':2,
    'reg_sec':4,
    'reg_lines':5,'reg_subs':6,'reg_trans':7,'reg_scada':8,'reg_other':9,'reg_tot':10,
    'acc_sec':12,
    'acc_lines':13,'acc_subs':14,'acc_trans':15,'acc_scada':16,'acc_other':17,'acc_tot':18,
    'delta_sec':20,'delta_reg_acc':21,'delta_reg_tax':22,
    'tab_sec':24,'tab_open':25,'tab_cap':26,'tab_dep':27,'tab_close':28,
    'accum_sec':30,'accum_reg':31,'accum_tax_open':32,'accum_tax_dep':33,'accum_tax_close':34,
    'rab_tab_diff':35,
    'dtl_sec':37,'dtl_open':38,'dtl_move':39,'dtl_close':40,
    'recon_sec':42,'recon_taxable':43,'recon_gross':44,'recon_gamma':45,
    'recon_net':46,'recon_check':47})

# ── Historical pre-computation ────────────────────────────────────
def _calc_rab_class_hist():
    """Per-class RAB roll-forward for FY2019A–FY2023A."""
    rab = {cls: op for cls,_,_,op,_ in RM_CLASSES}
    R = {}
    for idx, yr in enumerate(HIST):
        d = HIST_DATA[idx]; R[yr] = {}
        for cls, _, aer_life, _, alloc in RM_CLASSES:
            o = rab[cls]
            cpi_v = round(o * HV['cpi'][idx], 1)
            cap_v = round(d['cap'] * alloc, 1)
            dep_v = round(o / aer_life, 1)
            cl_v  = round(o + cpi_v + cap_v - dep_v, 1)
            R[yr][cls] = {'open':o,'cpi':cpi_v,'cap':cap_v,'dep':dep_v,'close':cl_v}
            rab[cls] = cl_v
    return R

RAB_CLASS_HIST = _calc_rab_class_hist()

def _calc_tab_class_hist():
    """Per-class TAB (Tax Asset Base) roll-forward for FY2019A–FY2023A."""
    tab = {cls: otab for cls,_,_,_,_,otab,_ in DK_CLASSES}
    R = {}
    for idx, yr in enumerate(HIST):
        d = HIST_DATA[idx]; R[yr] = {}
        for cls, _, ato_life, _, _, _, alloc in DK_CLASSES:
            o = tab[cls]
            cap_v  = round(d['cap'] * alloc, 1)
            dep_v  = round(o / ato_life, 1)
            cl_v   = round(o + cap_v - dep_v, 1)
            R[yr][cls] = {'open':o,'cap':cap_v,'dep':dep_v,'close':cl_v}
            tab[cls] = cl_v
    return R

TAB_CLASS_HIST = _calc_tab_class_hist()

# Pre-compute aggregate TAB totals for historical years
_CLS_KEYS = [c for c,_,_,_,_ in RM_CLASSES]
TAB_TOT_HIST = {}
_accum_tax = _OPEN_ACCUM_TAX_DEP
for _yr in HIST:
    _d = TAB_CLASS_HIST[_yr]
    _open  = round(sum(_d[c]['open']  for c in _CLS_KEYS), 1)
    _cap   = round(sum(_d[c]['cap']   for c in _CLS_KEYS), 1)
    _dep   = round(sum(_d[c]['dep']   for c in _CLS_KEYS), 1)
    _close = round(sum(_d[c]['close'] for c in _CLS_KEYS), 1)
    _accum_tax_open = _accum_tax
    _accum_tax += _dep
    TAB_TOT_HIST[_yr] = {'open':_open,'cap':_cap,'dep':_dep,'close':_close,
                         'accum_open':round(_accum_tax_open,1),
                         'accum_close':round(_accum_tax,1)}

# ── RAB Roll-Forward (Detailed) ────────────────────────────────────
def build_rab_rollforward(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws, 50, 11); frz(ws)
    title_row(ws, "RAB ROLL-FORWARD MODEL — Detailed by Asset Class (AER Standard Lives)", "4B0082")
    yr_hdr(ws, RM['yr_hdr'])

    # Per-class roll-forward
    for cls, label, aer_life, _, alloc in RM_CLASSES:
        sec(ws, RM[f'{cls}_sec'], label)
        for lbl_t, key in [
            (f"    Opening RAB ($M)", f'{cls}_open'),
            (f"    + CPI Indexation", f'{cls}_cpi'),
            (f"    + Capex Additions (×{int(alloc*100)}% allocation)", f'{cls}_cap'),
            (f"    − Regulatory Depreciation  (÷ {aer_life} yr AER life)", f'{cls}_dep'),
            (f"    = Closing RAB", f'{cls}_close'),
        ]:
            lbl(ws, RM[key], lbl_t)

        for i, yr in enumerate(ALL):
            col = DC+i; cl = get_column_letter(col); h = yr in HIST
            pv  = get_column_letter(col-1) if col > DC else None

            if h:
                cd = RAB_CLASS_HIST[yr][cls]
                hcell(ws, RM[f'{cls}_open'],  col, cd['open'])
                hcell(ws, RM[f'{cls}_cpi'],   col, cd['cpi'])
                hcell(ws, RM[f'{cls}_cap'],   col, cd['cap'])
                hcell(ws, RM[f'{cls}_dep'],   col, cd['dep'])
                hcell(ws, RM[f'{cls}_close'], col, cd['close'], bold=True)
            else:
                prev_yr_cl = get_column_letter(ycol(yr-1))
                fcell(ws, RM[f'{cls}_open'],  col,
                      f"={prev_yr_cl}{RM[f'{cls}_close']}")
                fcell(ws, RM[f'{cls}_cpi'],   col,
                      f"={cl}{RM[f'{cls}_open']}*Assumptions!{cl}{A['cpi']}")
                fcell(ws, RM[f'{cls}_cap'],   col,
                      xref("Fixed Assets", cl, FA[f'{cls}_adds']))
                fcell(ws, RM[f'{cls}_dep'],   col,
                      f"={cl}{RM[f'{cls}_open']}/{aer_life}")
                fcell(ws, RM[f'{cls}_close'], col,
                      f"={cl}{RM[f'{cls}_open']}+{cl}{RM[f'{cls}_cpi']}"
                      f"+{cl}{RM[f'{cls}_cap']}-{cl}{RM[f'{cls}_dep']}",
                      bold=True, tot=True)

    # Total section
    sec(ws, RM['tot_sec'], "TOTAL REGULATORY ASSET BASE ($M)")
    sub_map = {'tot_open':'open','tot_cpi':'cpi','tot_cap':'cap',
               'tot_dep':'dep','tot_close':'close'}
    tot_labels = {
        'tot_open':  "  Total Opening RAB",
        'tot_cpi':   "  Total CPI Indexation",
        'tot_cap':   "  Total Capex Additions",
        'tot_dep':   "  Total Regulatory Depreciation  →  Return of RAB (PTRM)",
        'tot_close': "  Total Closing RAB",
    }
    for rk, rl in tot_labels.items():
        lbl(ws, RM[rk], rl)
        sk = sub_map[rk]
        is_tot = (rk == 'tot_close')
        for i, yr in enumerate(ALL):
            col = DC+i; cl = get_column_letter(col); h = yr in HIST
            refs = "+".join(f"{cl}{RM[f'{c}_{sk}']}" for c in _CLS_KEYS)
            fcell(ws, RM[rk], col, f"={refs}",
                  bold=is_tot, tot=is_tot, hist=h)

    # Average RAB
    lbl(ws, RM['avg_rab'],
        "  Average RAB  [(Opening + Closing) ÷ 2]  →  Return on RAB base (PTRM)")
    for i, yr in enumerate(ALL):
        col = DC+i; cl = get_column_letter(col); h = yr in HIST
        fcell(ws, RM['avg_rab'], col,
              f"=({cl}{RM['tot_open']}+{cl}{RM['tot_close']})/2", hist=h)

    # Accumulated regulatory depreciation
    sec(ws, RM['accum_sec'], "Accumulated Regulatory Depreciation ($M)")
    lbl(ws, RM['accum_open'],  "  Opening Accumulated Reg. Dep.")
    lbl(ws, RM['accum_dep'],   "  Current Year Reg. Dep.")
    lbl(ws, RM['accum_close'], "  Closing Accumulated Reg. Dep.")
    for i, yr in enumerate(ALL):
        col = DC+i; cl = get_column_letter(col); h = yr in HIST
        pv  = get_column_letter(col-1) if col > DC else None
        if yr == 2019:
            fcell(ws, RM['accum_open'],  col, f"=Assumptions!$B${A['open_reg_dep']}", hist=True)
        else:
            fcell(ws, RM['accum_open'],  col, f"={pv}{RM['accum_close']}", hist=h)
        fcell(ws, RM['accum_dep'],   col, f"={cl}{RM['tot_dep']}", hist=h)
        fcell(ws, RM['accum_close'], col,
              f"={cl}{RM['accum_open']}+{cl}{RM['accum_dep']}",
              bold=True, tot=True, hist=h)

    # Regulatory period summaries — single merged row showing period totals
    sec(ws, RM['rcp_sec'], "Regulatory Period Summaries")
    rcp1_lets = [get_column_letter(ycol(yr)) for yr in range(2024,2029)]
    rcp2_lets = [get_column_letter(ycol(yr)) for yr in range(2029,2034)]
    summary_rows = [
        (RM['rcp1_cap'], "  RCP1 (FY2024–28)  Total Capex ($M)",          'tot_cap',   rcp1_lets),
        (RM['rcp1_dep'], "  RCP1 (FY2024–28)  Total Reg. Depreciation",   'tot_dep',   rcp1_lets),
        (RM['rcp1_avg'], "  RCP1 (FY2024–28)  Average Closing RAB",       'tot_close', rcp1_lets),
        (RM['rcp2_cap'], "  RCP2 (FY2029–33)  Total Capex ($M)",          'tot_cap',   rcp2_lets),
        (RM['rcp2_dep'], "  RCP2 (FY2029–33)  Total Reg. Depreciation",   'tot_dep',   rcp2_lets),
        (RM['rcp2_avg'], "  RCP2 (FY2029–33)  Average Closing RAB",       'tot_close', rcp2_lets),
    ]
    for row, rl, src_rk, lets in summary_rows:
        lbl(ws, row, rl)
        sum_expr = "+".join(f"{c}{RM[src_rk]}" for c in lets)
        if 'avg' in rl.lower():
            fcell(ws, row, DC, f"=({sum_expr})/{len(lets)}", sub=True)
        else:
            fcell(ws, row, DC, f"={sum_expr}", sub=True)

# ── AER Depreciation Tracking ──────────────────────────────────────
def build_dep_tracking(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws, 54, 11); frz(ws)
    title_row(ws,
        "AER DEPRECIATION TRACKING MODEL — Regulatory | Accounting | Tax (ATO)",
        "4B0082")
    yr_hdr(ws, DK['yr_hdr'])

    # ── Section A: Regulatory Depreciation ───────────────────────────
    sec(ws, DK['reg_sec'],
        "SECTION A — Regulatory Depreciation (AER Standard Lives) ($M)")
    reg_rows = [
        (DK['reg_lines'], "  Transmission Lines  (÷60yr AER)", 'lines_dep', 60),
        (DK['reg_subs'],  "  Substations  (÷55yr AER)",         'subs_dep',  55),
        (DK['reg_trans'], "  Transformers  (÷45yr AER)",         'trans_dep', 45),
        (DK['reg_scada'], "  SCADA & Control  (÷25yr AER)",      'scada_dep', 25),
        (DK['reg_other'], "  Other Equipment  (÷30yr AER)",      'other_dep', 30),
    ]
    for row, rl, rm_key, _ in reg_rows:
        lbl(ws, row, rl)
        for i, yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            fcell(ws, row, col,
                  xref("RAB Roll-Forward", cl, RM[rm_key]), hist=h)
    lbl(ws, DK['reg_tot'], "  Total Regulatory Depreciation  →  Return of RAB (PTRM)")
    for i, yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws, DK['reg_tot'], col,
              f"=SUM({cl}{DK['reg_lines']}:{cl}{DK['reg_other']})",
              bold=True, tot=True, hist=h)

    # ── Section B: Accounting Depreciation ───────────────────────────
    sec(ws, DK['acc_sec'],
        "SECTION B — Accounting Depreciation AASB 116 (from Fixed Assets tab) ($M)")
    acc_rows = [
        (DK['acc_lines'], "  Transmission Lines  (÷40yr acc.)", 'lines'),
        (DK['acc_subs'],  "  Substations  (÷40yr acc.)",         'subs'),
        (DK['acc_trans'], "  Transformers  (÷30yr acc.)",         'trans'),
        (DK['acc_scada'], "  SCADA & Control  (÷15yr acc.)",      'scada'),
        (DK['acc_other'], "  Other Equipment  (÷20yr acc.)",      'other'),
    ]
    for row, rl, cls in acc_rows:
        lbl(ws, row, rl)
        for i, yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            fcell(ws, row, col,
                  xref("Fixed Assets", cl, FA[f'{cls}_da_chg']), hist=h)
    lbl(ws, DK['acc_tot'],
        "  Total Accounting Depreciation  →  cross-check vs IS D&A charge")
    for i, yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws, DK['acc_tot'], col,
              f"=SUM({cl}{DK['acc_lines']}:{cl}{DK['acc_other']})",
              bold=True, tot=True, hist=h)

    # ── Section C: Depreciation Delta ────────────────────────────────
    sec(ws, DK['delta_sec'],
        "SECTION C — Depreciation Delta Analysis ($M)")
    lbl(ws, DK['delta_reg_acc'],
        "  Regulatory Dep − Accounting Dep  "
        "(>0 = reg slower → regulatory asset builds)")
    lbl(ws, DK['delta_reg_tax'],
        "  Regulatory Dep − Tax Dep  "
        "(>0 = reg slower than ATO → MAR includes more 'Return of RAB')")
    for i, yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        fcell(ws, DK['delta_reg_acc'], col,
              f"={cl}{DK['reg_tot']}-{cl}{DK['acc_tot']}", hist=h)
        fcell(ws, DK['delta_reg_tax'], col,
              f"={cl}{DK['reg_tot']}-{cl}{DK['tab_dep']}", hist=h)

    # ── Section D: Tax Asset Base (TAB) Roll-Forward ──────────────────
    sec(ws, DK['tab_sec'],
        "SECTION D — Tax Asset Base (TAB) Roll-Forward  "
        "[ATO Prime Cost | Blended Life: 37 years] ($M)")
    lbl(ws, DK['tab_open'],  "  Opening TAB")
    lbl(ws, DK['tab_cap'],   "  + Capex Additions  (ATO cost basis)")
    lbl(ws, DK['tab_dep'],
        f"  − Tax Depreciation  (Opening TAB ÷ {BLENDED_ATO_LIFE}yr blended ATO life)")
    lbl(ws, DK['tab_close'], "  = Closing TAB")
    for i, yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        pv = get_column_letter(col-1) if col > DC else None
        td = TAB_TOT_HIST.get(yr)
        if h and td:
            hcell(ws, DK['tab_open'],  col, td['open'])
            hcell(ws, DK['tab_cap'],   col, td['cap'])
            hcell(ws, DK['tab_dep'],   col, td['dep'])
            hcell(ws, DK['tab_close'], col, td['close'], bold=True)
        else:
            prev_yr_cl = get_column_letter(ycol(yr-1))
            fcell(ws, DK['tab_open'],  col,
                  f"={prev_yr_cl}{DK['tab_close']}")
            fcell(ws, DK['tab_cap'],   col,
                  xref("Capex", cl, CP['tot_capex']))
            fcell(ws, DK['tab_dep'],   col,
                  f"={cl}{DK['tab_open']}/{BLENDED_ATO_LIFE}")
            fcell(ws, DK['tab_close'], col,
                  f"={cl}{DK['tab_open']}+{cl}{DK['tab_cap']}-{cl}{DK['tab_dep']}",
                  bold=True, tot=True)

    # ── Section E: Accumulated Depreciation Comparison ───────────────
    sec(ws, DK['accum_sec'],
        "SECTION E — Accumulated Depreciation Comparison ($M)")
    lbl(ws, DK['accum_reg'],
        "  Accumulated Regulatory Dep. (closing)  ←  RAB Roll-Forward")
    lbl(ws, DK['accum_tax_open'],  "  Accumulated Tax Dep. (opening)")
    lbl(ws, DK['accum_tax_dep'],   "  + Current Year Tax Dep.")
    lbl(ws, DK['accum_tax_close'], "  = Accumulated Tax Dep. (closing)")
    lbl(ws, DK['rab_tab_diff'],
        "  Total RAB − Total TAB  (AER vs ATO asset base gap)")

    _accum_tax_running = _OPEN_ACCUM_TAX_DEP
    for i, yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        pv = get_column_letter(col-1) if col > DC else None
        # Accumulated reg dep from RAB Roll-Forward tab
        fcell(ws, DK['accum_reg'], col,
              xref("RAB Roll-Forward", cl, RM['accum_close']), hist=h)
        td = TAB_TOT_HIST.get(yr)
        if h and td:
            hcell(ws, DK['accum_tax_open'],  col, td['accum_open'])
            hcell(ws, DK['accum_tax_dep'],   col, td['dep'])
            hcell(ws, DK['accum_tax_close'], col, td['accum_close'], bold=True)
        else:
            prev_yr_cl = get_column_letter(ycol(yr-1))
            fcell(ws, DK['accum_tax_open'],  col, f"={prev_yr_cl}{DK['accum_tax_close']}")
            fcell(ws, DK['accum_tax_dep'],   col, f"={cl}{DK['tab_dep']}")
            fcell(ws, DK['accum_tax_close'], col,
                  f"={cl}{DK['accum_tax_open']}+{cl}{DK['accum_tax_dep']}",
                  bold=True, tot=True)
        # RAB − TAB
        fcell(ws, DK['rab_tab_diff'], col,
              xref("RAB Roll-Forward", cl, RM['tot_close'])
              .replace("=","=") + f"-{cl}{DK['tab_close']}",
              sub=True, hist=h)

    # ── Section F: Deferred Tax Liability (DTL) Roll-Forward ──────────
    sec(ws, DK['dtl_sec'],
        "SECTION F — Deferred Tax Liability (DTL) Roll-Forward ($M)  "
        "[AASB 112 | 30% corporate rate]")
    lbl(ws, DK['dtl_open'],  "  Opening DTL")
    lbl(ws, DK['dtl_move'],
        "  + Movement: (Tax Dep − Acc Dep) × 30%  "
        "(positive = DTL grows as ATO deductions exceed accounting)")
    lbl(ws, DK['dtl_close'], "  = Closing DTL  →  feeds Balance Sheet deferred tax")
    _dtl_open = _OPEN_DTL
    for i, yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        pv = get_column_letter(col-1) if col > DC else None
        if yr == 2019:
            hcell(ws, DK['dtl_open'], col, _dtl_open)
        elif h:
            fcell(ws, DK['dtl_open'], col, f"={pv}{DK['dtl_close']}", hist=True)
        else:
            fcell(ws, DK['dtl_open'], col, f"={pv}{DK['dtl_close']}")
        td = TAB_TOT_HIST.get(yr)
        if h and td:
            move = round((td['dep'] - HIST_DATA[i]['da']) * 0.30, 1)
            hcell(ws, DK['dtl_move'], col, move)
            _dtl_cl = round(_dtl_open + move, 1)
            hcell(ws, DK['dtl_close'], col, _dtl_cl, bold=True)
            _dtl_open = _dtl_cl
        else:
            fcell(ws, DK['dtl_move'], col,
                  f"=({cl}{DK['tab_dep']}-{cl}{DK['acc_tot']})*Assumptions!$B${A['tax_rate']}")
            fcell(ws, DK['dtl_close'], col,
                  f"={cl}{DK['dtl_open']}+{cl}{DK['dtl_move']}",
                  bold=True, tot=True)

    # ── Section G: PTRM Tax Allowance Reconciliation ──────────────────
    sec(ws, DK['recon_sec'],
        "SECTION G — PTRM Tax Allowance Reconciliation ($M)")
    lbl(ws, DK['recon_taxable'],
        "  Regulatory Taxable Income  (Return on RAB = Avg RAB × WACC)")
    lbl(ws, DK['recon_gross'],   "  Gross Tax  (× 30% corporate rate)")
    lbl(ws, DK['recon_gamma'],
        "  Less: Gamma Benefit  (× gamma × 30%)  AER imputation credit offset")
    lbl(ws, DK['recon_net'],     "  Net PTRM Tax Allowance  (Gross − Gamma)")
    lbl(ws, DK['recon_check'],   "  Cross-Check vs PTRM tab  (= 0 if matched)")
    for i, yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        if h:
            d = HIST_DATA[i]
            avg_rab  = d['avg_rab']
            taxable  = round(avg_rab * AV['wacc'], 1)
            gross    = round(taxable * AV['tax_rate'], 1)
            gamma_b  = round(taxable * AV['gamma'] * AV['tax_rate'], 1)
            net_t    = round(gross - gamma_b, 1)
            hcell(ws, DK['recon_taxable'], col, taxable)
            hcell(ws, DK['recon_gross'],   col, gross)
            hcell(ws, DK['recon_gamma'],   col, gamma_b)
            hcell(ws, DK['recon_net'],     col, net_t, bold=True)
            hcell(ws, DK['recon_check'],   col, 0)
        else:
            fcell(ws, DK['recon_taxable'], col,
                  xref("RAB Roll-Forward", cl, RM['avg_rab']))
            fcell(ws, DK['recon_gross'], col,
                  f"={cl}{DK['recon_taxable']}*Assumptions!$B${A['tax_rate']}")
            fcell(ws, DK['recon_gamma'], col,
                  f"={cl}{DK['recon_taxable']}*Assumptions!$B${A['gamma']}"
                  f"*Assumptions!$B${A['tax_rate']}")
            fcell(ws, DK['recon_net'], col,
                  f"={cl}{DK['recon_gross']}-{cl}{DK['recon_gamma']}",
                  bold=True, tot=True)
            chk_expr = (f"={cl}{DK['recon_net']}"
                        + xref("PTRM", cl, PT['tax_all']).replace("=", "-"))
            fcell(ws, DK['recon_check'], col, chk_expr, sub=True)
            # Colour check cell green/red
            c_obj = ws.cell(row=DK['recon_check'], column=col)
            c_obj.fill = C_OK

# ── Updated build() — 19 tabs ─────────────────────────────────────
def build():
    wb = Workbook()
    ws_cover  = wb.active;               ws_cover.title = "Cover"
    ws_assm   = wb.create_sheet("Assumptions")
    ws_dtm    = wb.create_sheet("DTM")
    ws_proj   = wb.create_sheet("Projects")
    ws_capex  = wb.create_sheet("Capex")
    ws_fa     = wb.create_sheet("Fixed Assets")
    ws_wf     = wb.create_sheet("Workforce")
    ws_mn     = wb.create_sheet("Maintenance")
    ws_refm   = wb.create_sheet("REFM")
    ws_front  = wb.create_sheet("Frontier")
    ws_rab    = wb.create_sheet("RAB")
    ws_rabrf  = wb.create_sheet("RAB Roll-Forward")
    ws_ptrm   = wb.create_sheet("PTRM")
    ws_deptrk = wb.create_sheet("Dep Tracking")
    ws_init   = wb.create_sheet("Initiatives")
    ws_is     = wb.create_sheet("Income Statement")
    ws_bs     = wb.create_sheet("Balance Sheet")
    ws_cf     = wb.create_sheet("Cash Flow")
    ws_chk    = wb.create_sheet("Checks")

    build_cover(ws_cover)
    build_assumptions(ws_assm)
    build_dtm(ws_dtm)
    build_projects(ws_proj)
    build_capex(ws_capex)
    build_fixed_assets(ws_fa)
    build_workforce(ws_wf)
    build_maintenance(ws_mn)
    build_refm(ws_refm)
    build_frontier(ws_front)
    build_rab(ws_rab)
    build_rab_rollforward(ws_rabrf)
    build_ptrm(ws_ptrm)
    build_dep_tracking(ws_deptrk)
    build_initiatives(ws_init)
    build_is(ws_is)
    build_bs(ws_bs)
    build_cf(ws_cf)
    build_checks(ws_chk)

    tab_colours = {
        "Cover":             "1F4E79",
        "Assumptions":       "375623",
        "DTM":               "7030A0",
        "Projects":          "BF8F00",
        "Capex":             "BF8F00",
        "Fixed Assets":      "BF8F00",
        "Workforce":         "BF8F00",
        "Maintenance":       "BF8F00",
        "REFM":              "7030A0",
        "Frontier":          "7030A0",
        "RAB":               "7030A0",
        "RAB Roll-Forward":  "4B0082",
        "PTRM":              "7030A0",
        "Dep Tracking":      "4B0082",
        "Initiatives":       "C00000",
        "Income Statement":  "C00000",
        "Balance Sheet":     "C00000",
        "Cash Flow":         "C00000",
        "Checks":            "808080",
    }
    for ws in wb.worksheets:
        colour = tab_colours.get(ws.title)
        if colour:
            ws.sheet_properties.tabColor = colour

    path = "three_way_finance_model.xlsx"
    wb.save(path)
    print(f"Saved: {path}  ({len(wb.worksheets)} sheets)")

# ════════════════════════════════════════════════════════════════
# CAPITAL PLANNING v1.2 INTEGRATION
# 22 AER asset classes | Project allocation matrix | Commissioning/CWIP
# As-incurred vs as-commissioned capex | Vintage-based depreciation
# ════════════════════════════════════════════════════════════════

# ── 22 AER asset class taxonomy ─────────────────────────────────
# (key, name, useful_life_yrs, legacy_5class_rollup)
AC_CLASSES = [
    ('ac1',  'Transmission Lines >132kV',     60, 'lines'),
    ('ac2',  'Transmission Lines <=132kV',    50, 'lines'),
    ('ac3',  'Underground Cables',            45, 'lines'),
    ('ac4',  'Power Transformers',            45, 'trans'),
    ('ac5',  'Switchgear',                    45, 'subs'),
    ('ac6',  'Substation Plant',              50, 'subs'),
    ('ac7',  'Protection & Control',          15, 'scada'),
    ('ac8',  'SCADA',                         10, 'scada'),
    ('ac9',  'Metering',                      15, 'scada'),
    ('ac10', 'Buildings',                     40, 'other'),
    ('ac11', 'Land',                         999, 'other'),
    ('ac12', 'IT Systems',                     5, 'scada'),
    ('ac13', 'Battery Storage',               15, 'other'),
    ('ac14', 'Renewable Connection Assets',   40, 'lines'),
    ('ac15', 'Mobile Plant',                  10, 'other'),
] + [(f'ac{i}', f'Placeholder AC{i}', 30, 'other') for i in range(16, 23)]

AC_KEYS      = [a for a,_,_,_ in AC_CLASSES]
AC_LIFE      = {a: l for a,_,l,_ in AC_CLASSES}
AC_NAME      = {a: n for a,n,_,_ in AC_CLASSES}
CLASS_ROLLUP = {a: r for a,_,_,r in AC_CLASSES}
LEGACY_KEYS  = ['lines','subs','trans','scada','other']
PIDS         = [f'p{n:03d}' for n in range(1, 11)]

# ── Enhanced project register ───────────────────────────────────
# (pid, name, program, budget, opening_cwip, comm_fy, pct_reg, pct_nonreg, pct_trans,
#   alloc {ac:%}, spend {fy:$M})
PROJECTS_V2 = [
    ('p001', 'P001 Grid Modernisation Program',           'AUG',   630, 15, 2028, 0.85, 0.10, 0.05,
        {'ac1':0.40,'ac5':0.20,'ac6':0.10,'ac7':0.10,'ac8':0.10,'ac12':0.10},
        {2024:90,2025:130,2026:150,2027:120,2028:80,2029:60}),
    ('p002', 'P002 Substation Upgrades — Zone A',         'REPL',  210,  5, 2026, 1.00, 0.00, 0.00,
        {'ac5':0.40,'ac6':0.30,'ac4':0.20,'ac7':0.10},
        {2024:70,2025:90,2026:50}),
    ('p003', 'P003 HVDC Interconnect (Stage 1)',          'AUG',   480,  0, 2029, 0.90, 0.00, 0.10,
        {'ac1':0.50,'ac3':0.20,'ac4':0.20,'ac5':0.05,'ac8':0.05},
        {2025:60,2026:120,2027:150,2028:120,2029:30}),
    ('p004', 'P004 New Connection — Wind Farm A',         'CONN',   45,  0, 2025, 1.00, 0.00, 0.00,
        {'ac14':0.60,'ac1':0.20,'ac5':0.10,'ac9':0.10},
        {2024:25,2025:20}),
    ('p005', 'P005 Transformer Replacement Program',      'REPL',  350,  0, 2031, 1.00, 0.00, 0.00,
        {'ac4':0.80,'ac5':0.10,'ac7':0.10},
        {2024:35,2025:40,2026:45,2027:50,2028:55,2029:55,2030:50,2031:20}),
    ('p006', 'P006 Protection Relay Upgrade',             'REPL',  120,  0, 2027, 1.00, 0.00, 0.00,
        {'ac7':0.70,'ac8':0.20,'ac9':0.10},
        {2024:30,2025:35,2026:30,2027:25}),
    ('p007', 'P007 New Connection — Solar Farm B',        'CONN',   80,  0, 2026, 1.00, 0.00, 0.00,
        {'ac14':0.55,'ac1':0.25,'ac5':0.10,'ac9':0.10},
        {2025:40,2026:40}),
    ('p008', 'P008 Transmission Reconductoring',          'REPL',  420,  0, 2031, 1.00, 0.00, 0.00,
        {'ac1':0.60,'ac2':0.30,'ac5':0.10},
        {2026:50,2027:70,2028:90,2029:90,2030:80,2031:40}),
    ('p009', 'P009 Digital Twin Infrastructure',          'AUG',    90,  0, 2029, 0.40, 0.20, 0.40,
        {'ac12':0.70,'ac8':0.20,'ac10':0.10},
        {2027:30,2028:35,2029:25}),
    ('p010', 'P010 Reactive Power Support',               'AUG',    75,  0, 2030, 1.00, 0.00, 0.00,
        {'ac6':0.50,'ac4':0.30,'ac5':0.20},
        {2028:25,2029:30,2030:20}),
]

# Validate allocation %
for _pid, _n, _, _, _, _, _, _, _, _alloc, _ in PROJECTS_V2:
    _s = sum(_alloc.values())
    assert abs(_s - 1.0) < 0.001, f"Project {_pid} allocation sums to {_s}, must be 1.0"

# ── Row maps for new tabs ───────────────────────────────────────
PRV2 = {}  # Project Register v2
CM = {}    # Commissioning
RO = {}    # Reg Output - Commissioned
ND = {}    # New Assets Depn
DS = {}    # Depn Summary

# PRV2: 10 projects × 30-row blocks (sec/attrs/22acs/3split/commyr/spend)
# Block layout: b=sec, b+1=attrs, b+2..b+23=ac1..ac22, b+24=reg, b+25=nonreg,
#   b+26=trans, b+27=commyr, b+28=spend, b+29=blank
PRV2.update({'title':1,'yr_hdr':2})
for _idx, _pid in enumerate(PIDS):
    _b = 4 + _idx * 30   # bases at 4, 34, 64, 94, 124, 154, 184, 214, 244, 274
    PRV2[f'{_pid}_sec']    = _b
    PRV2[f'{_pid}_attrs']  = _b + 1
    for _i, _ac in enumerate(AC_KEYS):
        PRV2[f'{_pid}_{_ac}'] = _b + 2 + _i  # rows b+2..b+23
    PRV2[f'{_pid}_reg']    = _b + 24
    PRV2[f'{_pid}_nonreg'] = _b + 25
    PRV2[f'{_pid}_trans']  = _b + 26
    PRV2[f'{_pid}_commyr'] = _b + 27
    PRV2[f'{_pid}_spend']  = _b + 28
PRV2.update({'tot_sec': 305, 'tot_spend': 306, 'tot_reg': 307,
             'ac_ref_sec': 310, 'ac_ref_hdr': 311})
for _i, _k in enumerate(AC_KEYS):
    PRV2[f'ac_{_k}'] = 312 + _i   # rows 312-333

# CM: 10 projects × 6 rows (sec/open/inc/comm/close/blank)
CM.update({'title':1,'yr_hdr':2,'hdr':4})
for _idx, _pid in enumerate(PIDS):
    _b = 5 + _idx * 6   # bases at 5, 11, 17, 23, 29, 35, 41, 47, 53, 59
    CM[f'{_pid}_sec']   = _b
    CM[f'{_pid}_open']  = _b + 1
    CM[f'{_pid}_inc']   = _b + 2
    CM[f'{_pid}_comm']  = _b + 3
    CM[f'{_pid}_close'] = _b + 4
CM.update({'tot_sec': 70, 'tot_open': 71, 'tot_inc': 72, 'tot_comm': 73, 'tot_close': 74})

# RO: 22 ACs at rows 5-26, total at 28
RO.update({'title':1,'yr_hdr':2,'hdr':4,'tot':28})
for _i, _k in enumerate(AC_KEYS):
    RO[_k] = 5 + _i

# ND: 22 ACs × 12-row blocks (sec + 10 vintage rows + subtotal)
ND.update({'title':1,'yr_hdr':2})
for _i, _ac in enumerate(AC_KEYS):
    _b = 4 + _i * 12   # bases at 4, 16, 28, ..., 256
    ND[f'{_ac}_sec'] = _b
    for _vi, _vyr in enumerate(range(2024, 2034)):
        ND[f'{_ac}_v{_vyr}'] = _b + 1 + _vi
    ND[f'{_ac}_tot'] = _b + 11
ND.update({'grand_sec': 268, 'grand_tot': 269})

# DS: New + Existing → totals by 5-class
DS.update({'title':1,'yr_hdr':2,
    'new_sec':4,'new_lines':5,'new_subs':6,'new_trans':7,'new_scada':8,'new_other':9,'new_tot':10,
    'exist_sec':12,'exist_lines':13,'exist_subs':14,'exist_trans':15,'exist_scada':16,
    'exist_other':17,'exist_tot':18,
    'tot_sec':20,'tot_lines':21,'tot_subs':22,'tot_trans':23,'tot_scada':24,
    'tot_other':25,'tot_reg_dep':26})


# ── Project Register v2 builder ──────────────────────────────────
def build_project_register_v2(ws):
    ws.sheet_view.showGridLines = False
    col_wid(ws, 50, 11); frz(ws)
    title_row(ws,
        "PROJECT REGISTER v2 — 22-AC Allocation Matrix | Reg Split | Commission Date",
        "BF8F00")
    yr_hdr(ws, PRV2['yr_hdr'])

    PROG_COLOR = {'AUG':'DCE6F1','REPL':'FFF2CC','CONN':'E2EFDA',
                  'TRANSFORM':'F4CCCC','NON-NETWORK':'EAD1DC'}

    for tup in PROJECTS_V2:
        pid, name, prog, budget, open_cwip, comm_fy, p_reg, p_nr, p_tr, alloc, spend = tup
        b = PRV2[f'{pid}_sec']

        # Section header row: project name, program, budget
        ws.cell(row=b, column=1,
                value=f"{name}  |  {prog}  |  Budget: ${budget}M").font = F_S
        ws.cell(row=b, column=1).fill = fill(PROG_COLOR.get(prog, 'FFFFFF'))
        for col in range(DC, DC+NT):
            ws.cell(row=b, column=col).fill = C_SEC

        # Attributes row
        ws.cell(row=PRV2[f'{pid}_attrs'], column=1,
                value="    Attributes:  Opening CWIP / Commission FY / %Reg / %Non-Reg / %Trans"
                ).font = mkf(italic=True, color="595959")

        # 22 AC allocation rows
        for ac in AC_KEYS:
            r = PRV2[f'{pid}_{ac}']
            lbl(ws, r, f"      {ac.upper()}  {AC_NAME[ac]}", indent=0)
            pct = alloc.get(ac, 0.0)
            inp(ws, r, DC, pct, FMTP)
            # Validation total in next column
            if ac == AC_KEYS[-1]:
                # last AC row: show sum check
                ws.cell(row=r, column=DC+1,
                        value=(f"=SUM(B{PRV2[f'{pid}_{AC_KEYS[0]}']}"
                               f":B{PRV2[f'{pid}_{AC_KEYS[-1]}']})")
                        ).number_format = FMTP
                ws.cell(row=r, column=DC+1).font = mkf(italic=True, color="1B6B1B")
                ws.cell(row=r, column=DC+1).alignment = RA

        # Reg split inputs
        lbl(ws, PRV2[f'{pid}_reg'],    "      % Regulated")
        inp(ws, PRV2[f'{pid}_reg'],    DC, p_reg, FMTP)
        lbl(ws, PRV2[f'{pid}_nonreg'], "      % Non-Regulated")
        inp(ws, PRV2[f'{pid}_nonreg'], DC, p_nr,  FMTP)
        lbl(ws, PRV2[f'{pid}_trans'],  "      % Transformation")
        inp(ws, PRV2[f'{pid}_trans'],  DC, p_tr,  FMTP)

        # Commission FY input (single cell col B)
        lbl(ws, PRV2[f'{pid}_commyr'], "      Commission FY (CWIP flips to in-service)")
        inp(ws, PRV2[f'{pid}_commyr'], DC, comm_fy, FMTN)

        # Annual spend row
        lbl(ws, PRV2[f'{pid}_spend'], "    Annual Spend $M (Capex Incurred)")
        for i, yr in enumerate(ALL):
            col = DC + i; h = yr in HIST
            amt = spend.get(yr, 0)
            if h:
                hcell(ws, PRV2[f'{pid}_spend'], col, amt, FMTD)
            else:
                fcell(ws, PRV2[f'{pid}_spend'], col, amt, FMTD, bold=True)

    # Aggregate totals at the bottom
    ws.cell(row=PRV2['tot_sec'], column=1,
            value="TOTAL ACROSS ALL PROJECTS").font = F_S
    for col in range(1, DC+NT):
        ws.cell(row=PRV2['tot_sec'], column=col).fill = C_SEC

    lbl(ws, PRV2['tot_spend'], "  Total Capex Incurred ($M)")
    for i, yr in enumerate(ALL):
        col = DC + i; cl = get_column_letter(col); h = yr in HIST
        refs = "+".join(f"{cl}{PRV2[f'{p}_spend']}" for p in PIDS)
        fcell(ws, PRV2['tot_spend'], col, f"={refs}",
              FMTD, bold=True, tot=True, hist=h)

    lbl(ws, PRV2['tot_reg'], "  Total Regulated Capex ($M)")
    for i, yr in enumerate(ALL):
        col = DC + i; cl = get_column_letter(col); h = yr in HIST
        refs = "+".join(f"{cl}{PRV2[f'{p}_spend']}*$B${PRV2[f'{p}_reg']}" for p in PIDS)
        fcell(ws, PRV2['tot_reg'], col, f"={refs}",
              FMTD, sub=True, hist=h)

    # ── Asset class reference table ──
    r_sec = PRV2['ac_ref_sec']
    ws.cell(row=r_sec, column=1,
            value="ASSET CLASS REFERENCE — 22 AER Classes  |  Useful Life  |  Legacy Rollup"
            ).font = F_S
    for c in range(1, 6):
        ws.cell(row=r_sec, column=c).fill = fill("4B0082")
        ws.cell(row=r_sec, column=c).font = mkf(bold=True, color="FFFFFF")

    r_hdr = PRV2['ac_ref_hdr']
    for c, txt in enumerate(["AC #", "Asset Class Name", "Useful Life (yrs)",
                              "Depn Rate (SL %pa)", "Legacy Rollup"], start=1):
        cc = ws.cell(row=r_hdr, column=c, value=txt)
        cc.fill = C_HDR; cc.font = F_H; cc.alignment = CA

    for row_i, (ac, name, life, rollup) in enumerate(AC_CLASSES):
        r = PRV2[f'ac_{ac}']
        ws.cell(row=r, column=1, value=ac.upper()).font = F_B
        ws.cell(row=r, column=2, value=name).font = F_L
        ws.cell(row=r, column=3, value=life if life < 999 else "N/A").number_format = FMTN
        ws.cell(row=r, column=3).alignment = RA
        rate_cell = ws.cell(row=r, column=4)
        if life >= 100:
            rate_cell.value = "Non-dep."
            rate_cell.font = mkf(italic=True, color="595959")
        else:
            rate_cell.value = f"=1/C{r}"
            rate_cell.number_format = "0.0%"
        rate_cell.alignment = RA
        rollup_cell = ws.cell(row=r, column=5, value=rollup.title())
        rollup_cell.font = F_L
        if row_i % 2 == 0:
            for c in range(1, 6):
                ws.cell(row=r, column=c).fill = fill("F5F0FF")


# ── Commissioning builder (CWIP roll-forward) ───────────────────
def build_commissioning(ws):
    ws.sheet_view.showGridLines = False
    col_wid(ws, 50, 11); frz(ws)
    title_row(ws,
        "COMMISSIONING SCHEDULE — CWIP Roll-Forward (Incurred → Commissioned)",
        "BF8F00")
    yr_hdr(ws, CM['yr_hdr'])

    ws.cell(row=CM['hdr'], column=1,
            value="Per-Project CWIP Roll-Forward: Opening + Incurred − Commissioned = Closing"
            ).font = mkf(italic=True, color="595959")

    PROG_COLOR = {'AUG':'DCE6F1','REPL':'FFF2CC','CONN':'E2EFDA',
                  'TRANSFORM':'F4CCCC','NON-NETWORK':'EAD1DC'}

    for tup in PROJECTS_V2:
        pid, name, prog, _, open_cwip, comm_fy, *_ = tup
        b = CM[f'{pid}_sec']
        ws.cell(row=b, column=1,
                value=f"{name}  |  Commission FY{comm_fy}").font = F_S
        ws.cell(row=b, column=1).fill = fill(PROG_COLOR.get(prog, 'FFFFFF'))
        for col in range(DC, DC+NT):
            ws.cell(row=b, column=col).fill = C_SEC

        lbl(ws, CM[f'{pid}_open'],  "    Opening CWIP")
        lbl(ws, CM[f'{pid}_inc'],   "    + Capex Incurred (from Project Register v2)")
        lbl(ws, CM[f'{pid}_comm'],  "    − Capex Commissioned (flips to in-service in FY shown)")
        lbl(ws, CM[f'{pid}_close'], "    = Closing CWIP")

        commyr_cell = f"$B${PRV2[f'{pid}_commyr']}"

        for i, yr in enumerate(ALL):
            col = DC + i; cl = get_column_letter(col); h = yr in HIST
            pv = get_column_letter(col-1) if col > DC else None

            # Opening CWIP: FY2024 = opening_cwip from data; otherwise prior closing
            if yr == 2024:
                fcell(ws, CM[f'{pid}_open'], col, open_cwip, FMTD)
            elif yr == 2019:
                hcell(ws, CM[f'{pid}_open'], col, 0, FMTD)
            else:
                fcell(ws, CM[f'{pid}_open'], col,
                      f"={pv}{CM[f'{pid}_close']}", FMTD, hist=h)

            # Incurred = Project Register v2 spend
            if h:
                hcell(ws, CM[f'{pid}_inc'], col, 0, FMTD)
            else:
                fcell(ws, CM[f'{pid}_inc'], col,
                      f"='Project Register v2'!{cl}{PRV2[f'{pid}_spend']}", FMTD)

            # Commissioned: IF year == commyr THEN open + incurred ELSE 0
            if h:
                hcell(ws, CM[f'{pid}_comm'], col, 0, FMTD)
            else:
                fcell(ws, CM[f'{pid}_comm'], col,
                      f"=IF({yr}>={commyr_cell},{cl}{CM[f'{pid}_open']}+{cl}{CM[f'{pid}_inc']},0)",
                      FMTD)

            # Closing CWIP
            if h:
                hcell(ws, CM[f'{pid}_close'], col, 0, FMTD, bold=True)
            else:
                fcell(ws, CM[f'{pid}_close'], col,
                      f"={cl}{CM[f'{pid}_open']}+{cl}{CM[f'{pid}_inc']}-{cl}{CM[f'{pid}_comm']}",
                      FMTD, bold=True, tot=True)

    # Aggregate totals
    ws.cell(row=CM['tot_sec'], column=1,
            value="PORTFOLIO TOTALS").font = F_S
    for col in range(1, DC+NT):
        ws.cell(row=CM['tot_sec'], column=col).fill = C_SEC

    for tot_row, tot_lbl, sub in [
        (CM['tot_open'],  "  Total Opening CWIP",            'open'),
        (CM['tot_inc'],   "  Total Capex Incurred",          'inc'),
        (CM['tot_comm'],  "  Total Capex Commissioned",      'comm'),
        (CM['tot_close'], "  Total Closing CWIP",            'close'),
    ]:
        lbl(ws, tot_row, tot_lbl)
        for i, yr in enumerate(ALL):
            col = DC + i; cl = get_column_letter(col); h = yr in HIST
            refs = "+".join(f"{cl}{CM[f'{p}_{sub}']}" for p in PIDS)
            is_close = (sub == 'close')
            fcell(ws, tot_row, col, f"={refs}",
                  FMTD, bold=is_close, tot=is_close, hist=h)


# ── Reg Output - Commissioned (22 AC × 15yr) ────────────────────
def build_reg_output_commissioned(ws):
    ws.sheet_view.showGridLines = False
    col_wid(ws, 50, 11); frz(ws)
    title_row(ws,
        "REG OUTPUT — CAPEX (As Commissioned, by Asset Class)",
        "BF8F00")
    yr_hdr(ws, RO['yr_hdr'])

    ws.cell(row=RO['hdr'], column=1,
            value="Commissioned capex split across 22 ACs via project allocation matrix"
            ).font = mkf(italic=True, color="595959")

    for ac in AC_KEYS:
        r = RO[ac]
        lbl(ws, r, f"  {ac.upper()}  {AC_NAME[ac]}")
        for i, yr in enumerate(ALL):
            col = DC + i; cl = get_column_letter(col); h = yr in HIST
            if h:
                hcell(ws, r, col, 0, FMTD)
            else:
                # Sum across 10 projects: commissioned $ × allocation %
                terms = [
                    f"'Commissioning'!{cl}{CM[f'{p}_comm']}*"
                    f"'Project Register v2'!$B${PRV2[f'{p}_{ac}']}"
                    for p in PIDS
                ]
                fcell(ws, r, col, "=" + "+".join(terms), FMTD)

    # Total row
    lbl(ws, RO['tot'], "  TOTAL — All Asset Classes ($M Commissioned)")
    for i, yr in enumerate(ALL):
        col = DC + i; cl = get_column_letter(col); h = yr in HIST
        first_ac = AC_KEYS[0]; last_ac = AC_KEYS[-1]
        fcell(ws, RO['tot'], col,
              f"=SUM({cl}{RO[first_ac]}:{cl}{RO[last_ac]})",
              FMTD, bold=True, tot=True, hist=h)


# ── New Assets Depreciation (vintage matrix per AC) ─────────────
def build_new_assets_depn(ws):
    ws.sheet_view.showGridLines = False
    col_wid(ws, 50, 11); frz(ws)
    title_row(ws,
        "NEW ASSETS DEPRECIATION — Vintage Cohort × Asset Class (Straight-Line)",
        "4B0082")
    yr_hdr(ws, ND['yr_hdr'])

    for ac in AC_KEYS:
        life = AC_LIFE[ac]
        b = ND[f'{ac}_sec']
        ws.cell(row=b, column=1,
                value=f"{ac.upper()}  {AC_NAME[ac]}  |  Useful Life: {life}yr"
                ).font = F_S
        for col in range(1, DC+NT):
            ws.cell(row=b, column=col).fill = C_SEC

        # 10 vintage cohort rows (FY2024..FY2033 commissioning years)
        for vyr in range(2024, 2034):
            r = ND[f'{ac}_v{vyr}']
            lbl(ws, r, f"    Vintage FY{vyr}E")
            vyr_cl = ylet(vyr)
            for i, t_yr in enumerate(ALL):
                col = DC + i; cl = get_column_letter(col); h = t_yr in HIST
                if h:
                    hcell(ws, r, col, 0, FMTD)
                elif life >= 100:
                    fcell(ws, r, col, 0, FMTD)
                elif vyr <= t_yr < vyr + life:
                    # Active depreciation
                    fcell(ws, r, col,
                          f"='Reg Output - Commissioned'!{vyr_cl}{RO[ac]}/{life}",
                          FMTD)
                else:
                    fcell(ws, r, col, 0, FMTD)

        # AC subtotal row
        r_tot = ND[f'{ac}_tot']
        lbl(ws, r_tot, f"    {ac.upper()} Total Depn")
        first_v = ND[f'{ac}_v2024']; last_v = ND[f'{ac}_v2033']
        for i, t_yr in enumerate(ALL):
            col = DC + i; cl = get_column_letter(col); h = t_yr in HIST
            fcell(ws, r_tot, col,
                  f"=SUM({cl}{first_v}:{cl}{last_v})",
                  FMTD, bold=True, sub=True, hist=h)

    # Grand total
    ws.cell(row=ND['grand_sec'], column=1,
            value="GRAND TOTAL — All Asset Classes").font = F_S
    for col in range(1, DC+NT):
        ws.cell(row=ND['grand_sec'], column=col).fill = C_SEC
    lbl(ws, ND['grand_tot'], "  Total New Assets Depreciation ($M)")
    for i, t_yr in enumerate(ALL):
        col = DC + i; cl = get_column_letter(col); h = t_yr in HIST
        refs = "+".join(f"{cl}{ND[f'{ac}_tot']}" for ac in AC_KEYS)
        fcell(ws, ND['grand_tot'], col, f"={refs}",
              FMTD, bold=True, tot=True, hist=h)


# ── Depreciation Summary ────────────────────────────────────────
def build_depn_summary(ws):
    ws.sheet_view.showGridLines = False
    col_wid(ws, 50, 11); frz(ws)
    title_row(ws,
        "DEPRECIATION SUMMARY — New Vintage + Existing Legacy → 5-Class Totals",
        "4B0082")
    yr_hdr(ws, DS['yr_hdr'])

    # ── Section A: New asset depn by 5-class rollup ──
    sec(ws, DS['new_sec'],
        "SECTION A — New Asset Depreciation (Vintage SL, rolled up to 5 classes) ($M)")
    for legacy in LEGACY_KEYS:
        r = DS[f'new_{legacy}']
        members = [a for a, _, _, rr in AC_CLASSES if rr == legacy]
        lbl(ws, r, f"  {legacy.title()}")
        for i, yr in enumerate(ALL):
            col = DC + i; cl = get_column_letter(col); h = yr in HIST
            refs = "+".join(f"'New Assets Depn'!{cl}{ND[f'{a}_tot']}" for a in members)
            fcell(ws, r, col, f"={refs}", FMTD, hist=h)
    lbl(ws, DS['new_tot'], "  Total New Assets Depreciation")
    for i, yr in enumerate(ALL):
        col = DC + i; cl = get_column_letter(col); h = yr in HIST
        fcell(ws, DS['new_tot'], col,
              f"=SUM({cl}{DS['new_lines']}:{cl}{DS['new_other']})",
              FMTD, bold=True, tot=True, hist=h)

    # ── Section B: Existing depn (per-class reg dep from RAB Roll-Forward) ──
    sec(ws, DS['exist_sec'],
        "SECTION B — Existing Asset Depreciation (from RAB Roll-Forward) ($M)")
    for legacy in LEGACY_KEYS:
        r = DS[f'exist_{legacy}']
        lbl(ws, r, f"  {legacy.title()}")
        for i, yr in enumerate(ALL):
            col = DC + i; cl = get_column_letter(col); h = yr in HIST
            fcell(ws, r, col,
                  f"='RAB Roll-Forward'!{cl}{RM[f'{legacy}_dep']}",
                  FMTD, hist=h)
    lbl(ws, DS['exist_tot'], "  Total Existing Depreciation")
    for i, yr in enumerate(ALL):
        col = DC + i; cl = get_column_letter(col); h = yr in HIST
        fcell(ws, DS['exist_tot'], col,
              f"=SUM({cl}{DS['exist_lines']}:{cl}{DS['exist_other']})",
              FMTD, bold=True, tot=True, hist=h)

    # ── Section C: Total regulatory depn (new + existing) ──
    sec(ws, DS['tot_sec'],
        "SECTION C — TOTAL REGULATORY DEPRECIATION (New + Existing) ($M)")
    for legacy in LEGACY_KEYS:
        r = DS[f'tot_{legacy}']
        lbl(ws, r, f"  {legacy.title()}")
        for i, yr in enumerate(ALL):
            col = DC + i; cl = get_column_letter(col); h = yr in HIST
            fcell(ws, r, col,
                  f"={cl}{DS[f'new_{legacy}']}+{cl}{DS[f'exist_{legacy}']}",
                  FMTD, hist=h)
    lbl(ws, DS['tot_reg_dep'],
        "  TOTAL Regulatory Depreciation  (parallel view to RAB Roll-Forward)")
    for i, yr in enumerate(ALL):
        col = DC + i; cl = get_column_letter(col); h = yr in HIST
        fcell(ws, DS['tot_reg_dep'], col,
              f"=SUM({cl}{DS['tot_lines']}:{cl}{DS['tot_other']})",
              FMTD, bold=True, tot=True, hist=h)


# ── Re-wired build_fixed_assets ──
def build_fixed_assets(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws,46,11); frz(ws)
    title_row(ws,"FIXED ASSET REGISTER — PP&E Roll-Forward by Asset Class","4A4A8A")
    yr_hdr(ws, FA['yr_hdr'])
    sec(ws,FA['asset_sec'],"Gross PP&E and Accumulated D&A by Asset Class ($M)")

    CLASSES=[
        ('lines',  "TRANSMISSION LINES",   40, 3200, 1067, 0.50),
        ('subs',   "SUBSTATIONS",           40, 2400,  800, 0.30),
        ('trans',  "TRANSFORMERS",          30, 1200,  480, 0.10),
        ('scada',  "SCADA & CONTROL",       15,  600,  180, 0.07),
        ('other',  "OTHER EQUIPMENT",       20,  400,   73, 0.03),
    ]

    for cls,label,life,og,oda,alloc in CLASSES:
        sec_row=FA[f'{cls}_sec']
        ws.cell(row=sec_row,column=1,
                value=f"{label}  (Standard Life: {life} yrs)").font=F_S
        for c in range(1,DC+NT): ws.cell(row=sec_row,column=c).fill=C_SEC

        for sub_lbl,key in [("  Opening Gross PP&E",f'{cls}_ogross'),
            ("  + Additions (incurred basis × class %)",f'{cls}_adds'),
            ("  − Disposals",f'{cls}_disp'),
            ("  Closing Gross PP&E",f'{cls}_cgross'),
            ("  Opening Accum. D&A",f'{cls}_oda'),
            ("  D&A Charge (closing gross / life)",f'{cls}_da_chg'),
            ("  Closing Accum. D&A",f'{cls}_cda')]:
            lbl(ws,FA[key],sub_lbl)

        for i,yr in enumerate(ALL):
            col=DC+i; cl=get_column_letter(col); h=yr in HIST
            pv=get_column_letter(col-1) if col>DC else None

            if yr==2019:
                fcell(ws,FA[f'{cls}_ogross'],col,og,FMTD,hist=True)
                fcell(ws,FA[f'{cls}_oda'],   col,oda,FMTD,hist=True)
            else:
                fcell(ws,FA[f'{cls}_ogross'],col,
                      f"={pv}{FA[f'{cls}_cgross']}",FMTD,hist=h)
                fcell(ws,FA[f'{cls}_oda'],   col,
                      f"={pv}{FA[f'{cls}_cda']}",FMTD,hist=h)

            capex_src=f"='Capex'!{cl}{CP['tot_capex']}*{alloc}"
            fcell(ws,FA[f'{cls}_adds'],col,capex_src,FMTD,hist=h)
            fcell(ws,FA[f'{cls}_disp'],col,0,FMTD,hist=h)
            fcell(ws,FA[f'{cls}_cgross'],col,
                f"={cl}{FA[f'{cls}_ogross']}+{cl}{FA[f'{cls}_adds']}-{cl}{FA[f'{cls}_disp']}",
                FMTD,hist=h)
            fcell(ws,FA[f'{cls}_da_chg'],col,
                f"={cl}{FA[f'{cls}_cgross']}/{life}",FMTD,hist=h)
            fcell(ws,FA[f'{cls}_cda'],col,
                f"={cl}{FA[f'{cls}_oda']}+{cl}{FA[f'{cls}_da_chg']}",FMTD,hist=h)

    # Totals section
    sec(ws,FA['tot_sec'],"CONSOLIDATED TOTALS ($M)")
    for tot_lbl,tot_key,sub_key in [
        ("  Total Gross PP&E",FA['tot_gross'],f'{{}}_cgross'),
        ("  Total Accum. D&A",FA['tot_accum_da'],f'{{}}_cda'),
        ("  Total Net PP&E",  FA['tot_net_ppe'], None),
        ("  Total D&A Charge (feeds IS D&A)",FA['da_charge'],f'{{}}_da_chg')]:
        lbl(ws,tot_key,tot_lbl)
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        for tot_key,sk in [(FA['tot_gross'],'_cgross'),(FA['tot_accum_da'],'_cda'),
                           (FA['da_charge'],'_da_chg')]:
            refs="+".join(f"{cl}{FA[f'{c}{sk}']}" for c in _FA_CLASSES)
            fcell(ws,tot_key,col,f"={refs}",FMTD,bold=True,hist=h)
        fcell(ws,FA['tot_net_ppe'],col,
            f"={cl}{FA['tot_gross']}-{cl}{FA['tot_accum_da']}",FMTD,bold=True,tot=True,hist=h)


# ── Re-wired build_rab — capex additions from Reg Output - Commissioned ──
def build_rab(ws):
    ws.sheet_view.showGridLines=False; col_wid(ws); frz(ws)
    title_row(ws,"REGULATORY ASSET BASE (RAB) ROLL-FORWARD","7B3F00")
    yr_hdr(ws,RB['yr_hdr'])
    sec(ws,RB['sec'],"RAB Roll-Forward ($M)")
    for lbl_t,key in [("  Opening RAB",RB['open_rab']),
                      ("  + CPI Indexation",RB['cpi_idx']),
                      ("  + Capex Additions (commissioned)",RB['capex']),
                      ("  − Regulatory Depreciation",RB['reg_dep']),
                      ("  = Closing RAB",RB['close_rab'])]:
        lbl(ws,key,lbl_t)
    lbl(ws,RB['avg_rab'],"  Average RAB (Return on RAB base)")
    sec(ws,RB['accum_sec'],"Accumulated Regulatory Depreciation ($M)")
    lbl(ws,RB['accum_reg_dep'],"  Accum. Reg. Depreciation (closing)")
    for i,yr in enumerate(ALL):
        col=DC+i; cl=get_column_letter(col); h=yr in HIST
        pv=get_column_letter(col-1) if col>DC else None
        if yr==2019:
            fcell(ws,RB['open_rab'],col,
                  f"=Assumptions!$B${A['open_rab']}",FMTD,hist=True)
        else:
            fcell(ws,RB['open_rab'],col,
                  f"={pv}{RB['close_rab']}",FMTD,hist=h)
        fcell(ws,RB['cpi_idx'],col,
              f"={cl}{RB['open_rab']}*Assumptions!{cl}{A['cpi']}",FMTD,hist=h)
        # History keeps REFM source; forecast uses Reg Output - Commissioned
        if h:
            fcell(ws,RB['capex'],col,xref("REFM",cl,RF['tot_capex']),FMTD,hist=True)
        else:
            fcell(ws,RB['capex'],col,
                  xref("Reg Output - Commissioned",cl,RO['tot']),FMTD)
        fcell(ws,RB['reg_dep'],col,
              f"={cl}{RB['open_rab']}/Assumptions!$B${A['std_life']}",FMTD,hist=h)
        fcell(ws,RB['close_rab'],col,
            f"={cl}{RB['open_rab']}+{cl}{RB['cpi_idx']}+{cl}{RB['capex']}-{cl}{RB['reg_dep']}",
            FMTD,bold=True,tot=True,hist=h)
        fcell(ws,RB['avg_rab'],col,
              f"=({cl}{RB['open_rab']}+{cl}{RB['close_rab']})/2",FMTD,hist=h)
        if yr==2019:
            fcell(ws,RB['accum_reg_dep'],col,
                f"=Assumptions!$B${A['open_reg_dep']}+{cl}{RB['reg_dep']}",FMTD,hist=True)
        else:
            fcell(ws,RB['accum_reg_dep'],col,
                f"={pv}{RB['accum_reg_dep']}+{cl}{RB['reg_dep']}",FMTD,hist=h)


# ── Updated build() — 27 tabs (Capital Planning v1.2 integrated) ─────
def build():
    wb = Workbook()
    ws_cover  = wb.active;               ws_cover.title = "Cover"
    ws_assm   = wb.create_sheet("Assumptions")
    ws_dtm    = wb.create_sheet("DTM")
    # Capital Planning v1.2 block (5 new tabs)
    ws_prv2   = wb.create_sheet("Project Register v2")
    ws_cm     = wb.create_sheet("Commissioning")
    ws_ro     = wb.create_sheet("Reg Output - Commissioned")
    ws_nd     = wb.create_sheet("New Assets Depn")
    ws_ds     = wb.create_sheet("Depn Summary")
    # Existing tabs continue
    ws_proj   = wb.create_sheet("Projects")
    ws_capex  = wb.create_sheet("Capex")
    ws_fa     = wb.create_sheet("Fixed Assets")
    ws_wf     = wb.create_sheet("Workforce")
    ws_mn     = wb.create_sheet("Maintenance")
    ws_refm   = wb.create_sheet("REFM")
    ws_front  = wb.create_sheet("Frontier")
    ws_rab    = wb.create_sheet("RAB")
    ws_rabrf  = wb.create_sheet("RAB Roll-Forward")
    ws_ptrm   = wb.create_sheet("PTRM")
    ws_deptrk = wb.create_sheet("Dep Tracking")
    ws_init   = wb.create_sheet("Initiatives")
    ws_is     = wb.create_sheet("Income Statement")
    ws_bs     = wb.create_sheet("Balance Sheet")
    ws_cf     = wb.create_sheet("Cash Flow")
    ws_chk    = wb.create_sheet("Checks")

    # Builder calls — strict dependency order
    build_cover(ws_cover)
    build_assumptions(ws_assm)
    build_dtm(ws_dtm)
    # Capital Planning v1.2
    build_project_register_v2(ws_prv2)
    build_commissioning(ws_cm)
    build_reg_output_commissioned(ws_ro)
    build_new_assets_depn(ws_nd)
    # Existing flow
    build_projects(ws_proj)
    build_capex(ws_capex)
    build_fixed_assets(ws_fa)
    build_workforce(ws_wf)
    build_maintenance(ws_mn)
    build_refm(ws_refm)
    build_frontier(ws_front)
    build_rab(ws_rab)
    build_rab_rollforward(ws_rabrf)
    build_ptrm(ws_ptrm)
    build_dep_tracking(ws_deptrk)
    # Depn Summary depends on RAB Roll-Forward
    build_depn_summary(ws_ds)
    build_initiatives(ws_init)
    build_is(ws_is)
    build_bs(ws_bs)
    build_cf(ws_cf)
    build_checks(ws_chk)

    tab_colours = {
        "Cover":                      "1F4E79",
        "Assumptions":                "375623",
        "DTM":                        "7030A0",
        "Project Register v2":        "BF8F00",
        "Commissioning":              "BF8F00",
        "Reg Output - Commissioned":  "BF8F00",
        "New Assets Depn":            "4B0082",
        "Depn Summary":               "4B0082",
        "Projects":                   "BF8F00",
        "Capex":                      "BF8F00",
        "Fixed Assets":               "BF8F00",
        "Workforce":                  "BF8F00",
        "Maintenance":                "BF8F00",
        "REFM":                       "7030A0",
        "Frontier":                   "7030A0",
        "RAB":                        "7030A0",
        "RAB Roll-Forward":           "4B0082",
        "PTRM":                       "7030A0",
        "Dep Tracking":               "4B0082",
        "Initiatives":                "C00000",
        "Income Statement":           "C00000",
        "Balance Sheet":              "C00000",
        "Cash Flow":                  "C00000",
        "Checks":                     "808080",
    }
    for ws in wb.worksheets:
        colour = tab_colours.get(ws.title)
        if colour:
            ws.sheet_properties.tabColor = colour

    path = "three_way_finance_model.xlsx"
    wb.save(path)
    print(f"Saved: {path}  ({len(wb.worksheets)} sheets)")

if __name__ == "__main__":
    build()
