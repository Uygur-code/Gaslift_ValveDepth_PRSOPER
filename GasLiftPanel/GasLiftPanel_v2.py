import pandas as pd
import numpy as np
import panel as pn
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler
pn.extension("plotly")
df=pd.read_excel('Gas_Lift_Results_Analysis.xlsx',sheet_name='Sheet1')
msP1 = pn.widgets.MultiSelect(name='P1', value=[0.91,1,1.099], options=[0.91, 1,1.099])
msP2 = pn.widgets.MultiSelect(name='P2', value=[0.91,1,1.099], options=[0.91, 1,1.099])
msWC = pn.widgets.MultiSelect(name='WC', value=[0,25,50,75], options=[0,25,50,75])
msPgi = pn.widgets.MultiSelect(name='Pgi', value=[1300], options=[1300,1500])
msPI = pn.widgets.MultiSelect(name='PI', value=[1], options=[1,30,50,100])
msPr = pn.widgets.MultiSelect(name='Pr', value=[1000, 1250, 1600], options=[1000, 1250, 1600]) 
msChkSz = pn.widgets.MultiSelect(name='Chk_Sz', value=[0.75,1,2], options=[0.75,1,2])
msGasGrav = pn.widgets.MultiSelect(name='Gas_SGravty', value=[0.7], options=[0.6, 0.7, 0.8, 0.9])
msHEP = pn.widgets.MultiSelect(name='HEP', value=[1335], options=[1335,1370,1420,1470,1520], size=5)
slctViolinx = pn.widgets.Select(name='Violin x-axis:', options=[ 'P1', 'P2', 'WC', 'Pgi', 'PI', 'Pr',  'Chk_Sz', 'Gas_SGravty', 'HEP' ])
sldrViolinWidth =pn.widgets.FloatSlider(start=0,end=10, step=0.001,value=0.01)
chkVioin_width=pn.widgets.Checkbox(name="Use", value=False)
violin_box = pn.WidgetBox('Custom Violin Width', chkVioin_width, sldrViolinWidth)
vln_width=pn.widgets.IntSlider(start=0, end=1000, step=50, value=700)
vln_hight=pn.widgets.IntSlider(start=0, end=1000, step=50, value=1000)
vln_sz=pn.WidgetBox('Violin plot size', vln_width, vln_hight)

@pn.depends(msP1.param.value, msP2.param.value, msWC.param.value, msPgi.param.value, msPr.param.value, msPI.param.value, msChkSz.param.value, msGasGrav.param.value, msHEP.param.value)
def tablx(vP1, vP2, vWC , vPgi, vPr, vPI, vChkSz, vGasGrav, vHEP):
    df_nvalve1=df[(df['P1'].isin(vP1)) & (df['P2'].isin(vP2)) & (df['WC'].isin(vWC)) & (df['Pgi'].isin(vPgi)) & (df['Pr'].isin(vPr))& (df['PI'].isin(vPI)) & (df['Chk_Sz'].isin(vChkSz)) & (df['Gas_SGravty'].isin(vGasGrav)) & (df['HEP'].isin(vHEP))]
    df_nvalve=pd.DataFrame(df_nvalve1)
    tbl_valve_0 = pd.DataFrame(df_nvalve1['valveDepth[0]'])
    tbl_valve_1 = pd.DataFrame(df_nvalve1['valveDepth[1]'])
    tbl_valve_2 = pd.DataFrame(df_nvalve1['valveDepth[2]'])
    tbl_valve_3 = pd.DataFrame(df_nvalve1['valveDepth[3]'])
    tbl_0=tbl_valve_0.describe(include=[np.number])
    tbl_1=tbl_valve_1.describe(include=[np.number])
    tbl_2=tbl_valve_2.describe(include=[np.number])
    tbl_3=tbl_valve_3.describe(include=[np.number])
    tblx=pd.concat([tbl_0,tbl_1, tbl_2, tbl_3], axis=1)
    return pn.widgets.Tabulator(tblx)
  
@pn.depends(msP1.param.value, msP2.param.value, msWC.param.value, msPgi.param.value, msPr.param.value, msPI.param.value, msChkSz.param.value, msGasGrav.param.value, msHEP.param.value)
def histx(vP1, vP2, vWC , vPgi, vPr, vPI, vChkSz, vGasGrav, vHEP):
    df_nvalve1=df['NVALVE'][(df['P1'].isin(vP1)) & (df['P2'].isin(vP2)) & (df['WC'].isin(vWC)) & (df['Pgi'].isin(vPgi)) & (df['Pr'].isin(vPr))& (df['PI'].isin(vPI)) & (df['Chk_Sz'].isin(vChkSz)) & (df['Gas_SGravty'].isin(vGasGrav)) & (df['HEP'].isin(vHEP))]
    df_nvalve=pd.DataFrame(df_nvalve1)
    n_nvalve={}
    n_nvalve[0]=np.count_nonzero(df_nvalve == 1)
    n_nvalve[1]=np.count_nonzero(df_nvalve == 2)
    n_nvalve[2]=np.count_nonzero(df_nvalve == 3)
    n_nvalve[3]=np.count_nonzero(df_nvalve == 4)
    binx=[1,2,3,4]
    sx_nvalve=n_nvalve.values()
    s_nvalve = [str(num) for num in sx_nvalve]
    figh =  px.bar(x=binx, y=n_nvalve)
    figh.update_yaxes(title="عدد الحالات",side="right", mirror=True, ticks='outside', showline=True, linecolor='black', gridcolor='lightgrey')
    figh.update_xaxes(autorange="reversed",title="عدد الصمامات",type="category", mirror=True, ticks='outside', showline=True, linecolor='black', gridcolor='lightgrey')
    figh.update_layout(autosize=False, width=300, height=400, font=dict(family="Times New Roman",size=14), plot_bgcolor="white") 
    figh.update_traces(text= s_nvalve, textposition='outside')
    return figh
    
@pn.depends(vln_width.param.value,vln_hight.param.value, chkVioin_width.param.value,sldrViolinWidth.param.value, slctViolinx.param.value, msP1.param.value, msP2.param.value, msWC.param.value, msPgi.param.value, msPr.param.value, msPI.param.value, msChkSz.param.value, msGasGrav.param.value, msHEP.param.value)
def violin(vlnW, vlnH,vchkViolin,vVwidth, vX, vP1, vP2, vWC , vPgi, vPr, vPI, vChkSz, vGasGrav, vHEP):
   l_P1=[0.91, 1,1.099]
   l_P2=[0.91, 1,1.099]
   l_WC=[0,25,50,75]
   l_Pgi=[1300,1500]
   l_PI=[1,30,50,100]
   l_Pr=[1000, 1250, 1600]
   l_Chk_Sz=[0.75,1,2]
   l_Gas_SGravty=[0.6, 0.7, 0.8, 0.9]
   l_HEP=[1335,1370,1420,1470,1520]
   l_empty=0
   if vX == "P1":
             n_X=len(l_P1)
             l_x=l_P1
             s_x="م1"
   elif vX == "P2":
             n_X=len(l_P2)
             l_x=l_P2
             s_x="م2"
   elif vX == "WC":
             n_X=len(l_WC)
             l_x=l_WC
             s_x="(%) القاطع المائي"
   elif vX == "Pgi":
             n_X=len(l_Pgi)
             l_x=l_Pgi
             s_x="ضغط حقن الغاز  (با\ع2)"
   elif vX == "PI":
             n_X=len(l_PI)
             l_x=l_PI
             s_x="الدالة الانتاجية (ب\ي\با\ع2)"
   elif vX == "Pr":
             n_X=len(l_Pr)
             l_x=l_Pr
             s_x="الضغط الباطني الساكن (با\ع2)"
   elif vX == "Chk_Sz":
             n_X=len(l_Chk_Sz)
             l_x=l_Chk_Sz
             s_x="فتحة الخانق (عقدة)"
   elif vX == "Gas_SGravty":
             n_X=len(l_Gas_SGravty)
             l_x=l_Gas_SGravty
             s_x="الوزن النوعي للغاز"
   elif vX == "HEP":
             n_X=len(l_HEP)
             l_x=l_HEP
             s_x="قاعدة البطانة (متر)"
   else :
            l_empty=1           
   if l_empty==0:
      
      df_rslt=df[(df['P1'].isin(vP1)) & (df['P2'].isin(vP2)) & (df['WC'].isin(vWC)) & (df['Pgi'].isin(vPgi)) & (df['Pr'].isin(vPr))& (df['PI'].isin(vPI)) & (df['Chk_Sz'].isin(vChkSz)) & (df['Gas_SGravty'].isin(vGasGrav)) & (df['HEP'].isin(vHEP))]
      xv=pd.Series(df_rslt['NVALVE']).max()
      xxy = 4
      if xv == 2:
            xxy = 2
      elif xv == 3:
            xxy = 3
      elif xv == 4:
            xxy = 4

      fig = make_subplots(rows=xxy, cols=1 ,  shared_xaxes=True)
      for sn in l_x:
          fig.add_trace(go.Violin(x=df_rslt[vX][df_rslt[vX]== sn],
                                  y=df_rslt['valveDepth[0]'][df_rslt[vX]== sn],
                                  name= sn,
                                  box_visible=True,
                                  meanline_visible=True),
                                  row=1, col=1)
      for sn in l_x:
          fig.add_trace(go.Violin(x=df_rslt[vX][df_rslt[vX]== sn],
                                  y=df_rslt['valveDepth[1]'][df_rslt[vX]== sn],
                                  name= sn,
                                  box_visible=True,
                                  meanline_visible=True),
                                  row=2, col=1)
      if xxy >= 3:
       for sn in l_x:
          fig.add_trace(go.Violin(x=df_rslt[vX][df_rslt[vX]== sn],
                                  y=df_rslt['valveDepth[2]'][df_rslt[vX]== sn],
                                  name= sn,
                                  box_visible=True,
                                  meanline_visible=True),
                                  row=3, col=1)
      
      if xxy == 4:
       for sn in l_x:
          fig.add_trace(go.Violin(x=df_rslt[vX][df_rslt[vX]== sn],
                                  y=df_rslt['valveDepth[3]'][df_rslt[vX]== sn],
                                  name= sn,
                                  box_visible=True,
                                  meanline_visible=True),
                                  row=4, col=1)
      if vchkViolin:
         fig.update_traces(width=vVwidth)
      fig.update_traces(box_width=0.1,showlegend=False)
      fig.update_layout(autosize=False, width=vlnW, height=vlnH, font=dict(family="Times New Roman",size=14), plot_bgcolor="white")
      fig.update_yaxes(autorange="reversed",title="عمق الصمام (متر)",side="right", mirror=True, ticks='outside', showline=True, linecolor='black', gridcolor='lightgrey')
      fig.update_xaxes(autorange="reversed", title=s_x, title_standoff=1,type="category", mirror=True, ticks='outside', showline=True, linecolor='black', gridcolor='lightgrey')
      paneViolin = pn.pane.Plotly(fig)  
   return paneViolin

@pn.depends(msP1.param.value, msP2.param.value, msWC.param.value, msPgi.param.value, msPr.param.value, msPI.param.value, msChkSz.param.value, msGasGrav.param.value, msHEP.param.value)
def pca_biplot(vP1, vP2, vWC , vPgi, vPr, vPI, vChkSz, vGasGrav, vHEP):
    df_in=df[(df['P1'].isin(vP1)) & (df['P2'].isin(vP2)) & (df['WC'].isin(vWC)) & (df['Pgi'].isin(vPgi)) & (df['Pr'].isin(vPr))& (df['PI'].isin(vPI)) & (df['Chk_Sz'].isin(vChkSz)) & (df['Gas_SGravty'].isin(vGasGrav)) & (df['HEP'].isin(vHEP))]
    features = [ 'P1','P2','WC', 'Pgi', 'Pr', 'PI', 'Chk_Sz', 'Gas_SGravty', 'HEP']
    X = df_in[features]
    X = StandardScaler().fit_transform(X)
    pca= PCA(n_components=2,random_state=1234)
    components=pca.fit_transform(X)
    loadings = pca.components_.T * np.sqrt(pca.explained_variance_)
    fig = px.scatter(components, x=0, y=1, color=df_in['valveDepth[0]'])
    for i, feature in enumerate(features):
        fig.add_annotation(ax=0, ay=0, axref="x", ayref="y", x=loadings[i,0], y=loadings[i,1], showarrow=True, arrowsize=2, arrowhead=2, xanchor="right", yanchor="bottom")
        fig.add_annotation(x=loadings[i,0], y=loadings[i,1], ax=0, ay=0, xanchor="center", yanchor="bottom", text=feature, yshift=5)
    return pn.pane.Plotly(fig)    

template = pn.template.BootstrapTemplate(
    title='GasLiftPanel',
    sidebar=[violin_box, slctViolinx, msHEP, msPI, msPgi, msGasGrav, msP1, msP2, msWC, msPr, msChkSz ],  
)
template.main.append(pn.Row(pn.Column(vln_sz, pn.Card(violin,title='Violin')), pn.Card(tablx,title='Table'), pn.Card(histx,title='Histogram'),pn.Card(pca_biplot,title='PCA Biplot')))
template.show();
template.servable();
