import pandas as pd
import itertools as it
from openserver import OpenServer
from openpyxl import  load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils.datetime import to_excel
from openpyxl.chart import (ScatterChart, Reference, Series)
import gc
import timeit
c=OpenServer()
c.connect()
dp_Valve=100.0
df_portsize=pd.DataFrame()
df_Ranges=pd.read_excel('Prosper_Input_Ranges.XLSX',sheet_name='Input_ranges')
l_P1=df_Ranges['P1'].loc[df_Ranges['P1']>=0].values.tolist()
l_P2=df_Ranges['P2'].loc[df_Ranges['P2']>=0].values.tolist()
l_GOR=df_Ranges['GOR'].loc[df_Ranges['GOR']>=0].values.tolist()
l_WHFP=df_Ranges['WHFP'].loc[df_Ranges['WHFP']>=0].values.tolist()
l_WC=df_Ranges['WC'].loc[df_Ranges['WC']>=0].values.tolist()
l_Pgi=df_Ranges['Pgi'].loc[df_Ranges['Pgi']>=0].values.tolist()
l_PI=df_Ranges['PI'].loc[df_Ranges['PI']>=0].values.tolist()
l_Pr=df_Ranges['Pr'].loc[df_Ranges['Pr']>=0].values.tolist()
l_Chk_Sz=df_Ranges['Chk_Sz'].loc[df_Ranges['Chk_Sz']>=0].values.tolist()
l_Gas_SGravity=df_Ranges['Gas_SGravity'].loc[df_Ranges['Gas_SGravity']>=0].values.tolist()
l_HEP=df_Ranges['HEP'].loc[df_Ranges['HEP']>=0].values.tolist()
lp = it.product(l_P1,l_P2,l_GOR,l_WHFP,l_WC,l_Pgi,l_PI,l_Pr,l_Chk_Sz,l_Gas_SGravity,l_HEP)
del df_Ranges
del l_P1, l_P2, l_GOR, l_WHFP, l_WC, l_Pgi, l_PI, l_Pr, l_Chk_Sz, l_Gas_SGravity, l_HEP
gc.collect()
df_output=pd.DataFrame()        # Create empty dataframe  to add results
dl=1428.0     #DL in RTKB based on AverageRTKB=270 mASL
starttime=timeit.default_timer()  
i=0
for P1 ,P2,GOR, WHFP,WC,Pgi,PI, Pr,Chk_Sz, Gas_SGravty, HEP in lp:
    Pgix=Pgi
    c.DoCmd('PROSPER.RESET(OUT)')
    df_output.loc[i,'P1']=P1
    df_output.loc[i,'P2']=P2
    df_output.loc[i,'GOR']=GOR
    df_output.loc[i,'WHFP']=WHFP
    df_output.loc[i,'WC']=WC
    df_output.loc[i,'Pgi']=Pgix
    df_output.loc[i,'PI']=PI
    df_output.loc[i,'Pr']=Pr
    df_output.loc[i,'Chk_Sz']=Chk_Sz
    df_output.loc[i,'Gas_SGravty']=Gas_SGravty
    df_output.loc[i,'HEP']=HEP
    Pr_HEP=Pr+3.2808*0.324*(HEP+5.0-dl)   # Reservoir pressure at casing shoe
    c.DoSet('PROSPER.SIN.EQP.Down.Data[1].Depth',HEP)   # Tubing
    c.DoSet('PROSPER.SIN.EQP.Down.Data[2].Depth',HEP+5.0)   # Casing
    c.DoSet('PROSPER.SIN.EQP.Devn.Data[1].Md',HEP+5.0)
    c.DoSet('PROSPER.SIN.EQP.Devn.Data[1].Tvd',HEP+5.0)
    c.DoSet('PROSPER.SIN.EQP.Geo.Data[1].Md',HEP+5.0)
    c.DoSet('PROSPER.SIN.EQP.Surf.Data[1].ID',Chk_Sz)
    c.DoSet('PROSPER.SIN.GLF.Method',0)     # Fixed Depth of Injection
    c.DoSet('PROSPER.ANL.COR.Corr[{Petroleum Experts 3}].A[0]',P1)
    c.DoSet('PROSPER.ANL.COR.Corr[{Petroleum Experts 3}].A[1]',P2)
    # IPR data update
    c.DoSet('PROSPER.SIN.IPR.Single.Pindex',PI)
    c.DoSet('PROSPER.SIN.IPR.Single.Wc',WC)
    c.DoSet('PROSPER.SIN.IPR.Single.totgor',GOR)  
    c.DoSet('PROSPER.SIN.IPR.Single.Pres',Pr_HEP)
    c.DoCmd('PROSPER.IPR.CALC')                 # Perform the calculation cmd in the inflow section
    # Gaslift input data
    c.DoSet('PROSPER.SIN.GLF.Method',0)
    c.DoSet('PROSPER.SIN.GLF.Gravity' ,Gas_SGravty)    
    # Gaslift design - New well
    InjMaxDepth=HEP-5.0
    c.DoSet('PROSPER.ANL.GLD.DesignMethod',1)
    c.DoSet('PROSPER.ANL.GLD.TubingLabel',	'PetroleumExperts3')
    c.DoSet('PROSPER.ANL.GLD.PipeLabel','BeggsandBrill')
    c.DoSet('PROSPER.ANL.GLD.MaxProd',5000.0)     # Maxliquid rate
    c.DoSet('PROSPER.ANL.GLD.MaxGas',5.0)         # Max gas available
    c.DoSet('PROSPER.ANL.GLD.MaxGasUL',5.0)       # Max gas during uploading
    c.DoSet('PROSPER.ANL.GLD.Fwhp',WHFP)        # flowing top node pressure
    c.DoSet('PROSPER.ANL.GLD.Uwhp',WHFP)        # Unloading top node pressure
    c.DoSet('PROSPER.ANL.GLD.OpInjPres',Pgix)    # Operating injection pressure
    c.DoSet('PROSPER.ANL.GLD.KoInjPres',1300.0)    # Kickoff injection pressure
    c.DoSet('PROSPER.ANL.GLD.dPValve',dp_Valve)      # desired dP across valve
    c.DoSet('PROSPER.ANL.GLD.MaxDepth',InjMaxDepth)    # Max depth of injection
    c.DoSet('PROSPER.ANL.GLD.MinSpace',75.0)      # Min spacing (m)
    c.DoSet('PROSPER.ANL.GLD.StatGrad',0.43)    # Static gradient of load fluid
    c.DoSet('PROSPER.ANL.GLD.WC',WC)
    c.DoSet('PROSPER.ANL.GLD.Solgor',GOR)          # Total GOR
    c.DoSet('PROSPER.ANL.GLD.ValveD1',50.0)       # Min chp decrease per valve    
    c.DoCmd('PROSPER.ANL.GLN.CALC')             # Perform gaslift design (new well)  
    c.DoCmd('PROSPER.REFRESH')    
    # Design results
    df_output.loc[i,'liq_rate']=c.DoGet('PROSPER.OUT.GLN.Legend[4]')
    df_output.loc[i,'oil_rate']=c.DoGet('PROSPER.OUT.GLN.Legend[5]')
    df_output.loc[i,'gas_inj_rate']=c.DoGet('PROSPER.OUT.GLN.Legend[6]')
    df_output.loc[i,'inj_press']=c.DoGet('PROSPER.OUT.GLN.Legend[7]')
    n_valve=c.DoGet('PROSPER.OUT.GLN.NVALR')  # No of valves
    md_valve=c.DoGet('PROSPER.OUT.GLN.PTRMSD[$]')  # Measured Depth of VALVE
    port_size=c.DoGet('PROSPER.OUT.GLN.PORT[$]')
    md_valve=[im for im in md_valve if im !=0 and im > 20]
    port_size=[im for im in port_size if im !=0 ]
    i_valve=n_valve
    t=0
    [c.DoSet('PROSPER.SIN.GLF.Depth[' + str(t) +']',md_valve[t])  for t in range(i_valve)]
    df_output.loc[i,'NVALVE']=n_valve
    for n  in range(i_valve):
        df_portsize.loc[i,'portSize[' + str(n)+ ']' ]= port_size[n]    
        df_output.loc[i,'valveDepth[' + str(n)+ ']' ]= md_valve[n]    
    i=i+1       # loop increment to for P1 ,P2,freeGOR, WHFP,WC,Pgi,PI, Pr,Chk_Sz, Gas_SGravty in lp: 
df_output.to_excel('GAS_LIFT_RESULTS.xlsx') 
df_portsize.to_excel('PORT_SIZE.xlsx')        # Dataframes to Excel file
del df_output, df_portsize
del md_valve, port_size
del lp
gc.collect()
c.disconnect()
print(timeit.default_timer()-starttime)