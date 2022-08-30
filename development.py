import pandas as pd
import numpy as np
import xlrd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import msoffcrypto
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
from email.mime.image import MIMEImage
import os
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib.pyplot import figure
from matplotlib.ticker import MaxNLocator
import matplotlib.ticker as mtick
import matplotlib.ticker as ticker
from matplotlib import gridspec

server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
msg['To']='eunbi1.yoon@lge.com'

#Subject 꾸미기
today=date.today()
today=today.strftime('%m/%d')
msg['Subject']='[1주차 주간 보고] 법인별 재료비 추이 보고'

########################## TL 만들기 ################################
########################## sub plot 만들기
#그래프 만들기
fig, ax = plt.subplots(1,3)
fig = plt.figure(constrained_layout = True, figsize=(16,9))
gs = fig.add_gridspec(12,16)
ax[0] = fig.add_subplot(gs[0:2,0:16])
ax[1] = fig.add_subplot(gs[3:8,0:16])

# Chart data 읽기
data1=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='TL_1')
data1.index=data1['Changing reason']
data1=data1.drop(['Changing reason'],axis=1)
ax[0].set_axis_off()
col_Colours=np.full(12,'#F8E0F6')
row_Colours=['#F6CCC8','#F6CCC8','#F6CCC8','#D6D6D6']

table1 = ax[0].table(cellText=data1.values,
                      rowLabels=data1.index,colLabels=data1.columns,loc='left',colColours=col_Colours, rowColours=row_Colours)


table1.auto_set_font_size(False)
table1.set_fontsize(9)
table1.auto_set_column_width(col=list(range(len(data1.columns))))
ax[0].annotate('PCB Price 0.79↑, Damper friction ↓0.036 ',xy=(0.797,0.5),color='green',fontsize=9)
ax[0].annotate('TH decreased $0.3 (Exchange Rate 0.5↓, Material increase 0.2↑), TN increased $0.78 (Material Increase 0.78↑)',fontsize=11,xy=(0,1))


########################## sub plot 만들기
# data읽기
data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/Eunbi/Automatic Email Request/Development.xlsx',sheet_name='FL')
data=data.set_index('Tool')
table_data=data.T
chart_data=data

# No Percent data
np_data=table_data[['TN Victor2 Pro White','TN Victor2 Pro Silver']]
npp_data=table_data[['VH Victor2 Pro White','VH Victor2 Pro Silver']]
np_data=np_data.round(1)
npp_data=npp_data.round(1)

# Percent data
p_data=table_data[['Victor2 Pro White Diff Rate(%)']]
pp_data=table_data[['Victor2 Pro Silver Diff Rate(%)']]
p_data=p_data.round(2)
pp_data=pp_data.round(2)

# 데이터 넣기
axF1=np_data.plot(use_index=True, color=['#D2D6D1','#D2D6D1'],
                  marker='o',markersize=4,
                  label=['TN Victor2 Pro White','TN Victor2 Pro Silver'],
                  ax=ax[1],legend=None)
axF1=npp_data.plot(use_index=True, color=['#EE2113','#EE2113'],
                  marker='o',markersize=4,
                  label=['VH Victor2 Pro White','VH Victor2 Pro Silver'],
                  ax=ax[1],legend=None) 

axF2 = axF1.twinx()

p_data.plot(kind='line', use_index=True,legend=None,
            color='#7CF9A2',
            marker='o',markersize=4,
            label=['Victor2 Pro White Diff Rate(%)'],ax=axF2)

pp_data.plot(kind='line', use_index=True,legend=None,
            color='#B6F97C',
             marker='o',markersize=4,
            label=['Victor2 Pro Silver Diff Rate(%)'],ax=axF2)




# UI
x = np.arange(18)
y1=list(np.array(table_data['TN Victor2 Pro White'].round(1).tolist()))
y2=list(np.array(table_data['VH Victor2 Pro White'].round(1).tolist()))
y3=list(np.array(table_data['Victor2 Pro White Diff Amount'].round(1).tolist()))
y4=list(np.array((table_data['Victor2 Pro White Diff Rate(%)']).round(2).tolist()))
y5=list(np.array(table_data['TN Victor2 Pro Silver'].round(1).tolist()))
y6=list(np.array(table_data['VH Victor2 Pro Silver'].round(1).tolist()))
y7=list(np.array(table_data['Victor2 Pro Silver Diff Amount'].round(1).tolist()))
y8=list(np.array((table_data['Victor2 Pro Silver Diff Rate(%)']).round(2).tolist()))


a =[30.3, 31.3, 32.0, 32.7, 32.1,32.2,31.4,31.3,31.4,31.3,30.5,30.1,30.0,30.0,30.7,31.3,31.3,31.4,32.6,33.1,33.6,33.7]
b='kVND/$'
#
for i in range(18):
    axF1.annotate(y1[i],xy=(x[i],y1[i]-0.02),va='top',color='#B5A9A7',fontsize=9)
    axF1.annotate(y2[i],xy=(x[i],y2[i]-0.02),va='top',color='red',fontsize=9)
    axF2.annotate(str(y4[i])+"%",xy=(x[i],y4[i]-0.02),va='top',color='#00B050',fontsize=9)
    axF1.annotate(y5[i],xy=(x[i],y5[i]-0.02),va='bottom',color='#B5A9A7',fontsize=9)
    axF1.annotate(y6[i],xy=(x[i],y6[i]-0.02),va='top',color='red',fontsize=9)
    axF2.annotate(str(y8[i])+"%",xy=(x[i],y8[i]-0.02),va='bottom',color='#92D050',fontsize=9)
    plt.text(x[i]-0.4, 42 ,a[i],color='green')
    i=i+1

plt.text(-1.5, 42 , b,color='green')
    

plt.title("Material Comparsion with VH(Sep 2W)",fontsize=13, pad='30')
axF1.set_ylabel('USD($)',color='gray')
axF2.set_ylabel('Percentage(%)',color='gray')
axF2.yaxis.set_major_formatter(mtick.PercentFormatter()) # y축 퍼센트로 만들기
axF1.set_ylim(200,320)
axF2.set_ylim(-5,40)
axF2.set_xticks(np.arange(18))
axF2.set_xticklabels(list(p_data.index))
axF2.set_xlim([-0.5,17.5])
axF1.legend(loc='upper left')
axF2.legend(loc='upper right',bbox_to_anchor=(0.41,1))

# 차트 만들기
ax[1].set_axis_off()

array1=np.full(18,'#F8E0F6') # col 색깔 채우기
array2=['#D2D6D1','#F6CCC8','white','#A3E7B3','#D2D6D1','#F6CCC8','white','#D8FDC9'] # row 색깔 채우기

the_table = plt.table(cellText=chart_data.round(2).values,
                      rowLabels=chart_data.index,
                      colLabels=chart_data.columns,
                      colColours=array1,
                      rowColours=array2)


plt.subplots_adjust(left=0.17,bottom=-0.1)
the_table.auto_set_font_size(False)
the_table.set_fontsize(9)


plt.savefig('FL.png')



##################### FL 데이터 뽑기 #######################
# TN Pro BW
#TN_Pro_BW=pd.read_excel('//US-SO11-NA08765/R&D Secrets/VI/FL/TN Pro BW_1004.xlsx',sheet_name=0)
#TN_Pro_BW_cost=TN_Pro_BW.at[12,'Unnamed: 16']# total
#TN_Pro_BW_cost=round(TN_Pro_BW_cost,1)
#print(TN_Pro_BW_cost)

#TN_Pro_BW_Change=TN_Pro_BW[['Unnamed: 6','Unnamed: 18']]
#TN_Pro_BW_Change.columns=['Item','Cost']
#TN_Pro_BW_Change=TN_Pro_BW_Change.drop(10,axis=0)
#TN_Pro_BW_Change=TN_Pro_BW_Change.drop(11,axis=0)
#TN_Pro_BW_Change=TN_Pro_BW_Change.dropna()

#TN_Pro_BW_Change.reset_index(drop=True, inplace=True)
#print(TN_Pro_BW_Change)

#TN_Pro_BW_Before=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name=1)
#TN_Pro_BW_Before=round(TN_Pro_BW_Before.at[0,21.09],1)
#print(TN_Pro_BW_Before)


# TN Pro SS
#TN_Pro_SS=pd.read_excel('//US-SO11-NA08765/R&D Secrets/VI/FL/TN Pro SS_1004.xlsx',sheet_name=0)
#TN_Pro_SS_cost=TN_Pro_SS.at[12,'Unnamed: 16']# total
#TN_Pro_SS_cost=round(TN_Pro_SS_cost,1)
#print(TN_Pro_SS_cost)

#TN_Pro_SS_Before=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name=1)
#TN_Pro_SS_Before=round(TN_Pro_SS_Before.at[4,21.09],1)
#print(TN_Pro_SS_Before)


# TN Pro SS
#TN_Pro_SS=pd.read_excel('//US-SO11-NA08765/R&D Secrets/VI/FL/TN Pro SS_1004.xlsx',sheet_name=0)
#TN_Pro_SS_cost=TN_Pro_SS.at[12,'Unnamed: 16']# total
#TN_Pro_SS_cost=round(TN_Pro_SS_cost,1)
#print(TN_Pro_SS_cost)

#TN_Pro_SS_Before=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name=1)
#TN_Pro_SS_Before=round(TN_Pro_SS_Before.at[4,21.09],1)
#print(TN_Pro_SS_Before)




########################## FL 만들기 ################################
########################## sub plot 만들기
#그래프 만들기
fig, ax = plt.subplots(1,3)
fig = plt.figure(constrained_layout = True, figsize=(16,12))
gs = fig.add_gridspec(12,16)
ax[0] = fig.add_subplot(gs[0:2,0:16])
ax[1] = fig.add_subplot(gs[3:8,0:16])

# Chart data 읽기
data1=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='FL_1')
data1.index=data1['Changing reason']
data1=data1.drop(['Changing reason'],axis=1)
ax[0].set_axis_off()
col_Colours=np.full(12,'#F8E0F6')
row_Colours=['#F6CCC8','#F6CCC8','#F6CCC8','#D6D6D6']

table1 = ax[0].table(cellText=data1.values,
                      rowLabels=data1.index,colLabels=data1.columns,loc='left',colColours=col_Colours, rowColours=row_Colours)


table1.auto_set_font_size(False)
table1.set_fontsize(9)
table1.auto_set_column_width(col=list(range(len(data1.columns))))
ax[0].annotate('Screw,Tapping: $0.0018↑',xy=(0.797,0.5),color='green',fontsize=9)
ax[0].annotate('VH decrease by $0.02↓ (Exchange Rate: $0.02↓) and TN is increased by $0.16↑ (Exchange Rate: $0.16↑)',fontsize=11,xy=(0,1))


########################## sub plot 만들기
# data읽기
data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/Eunbi/Automatic Email Request/Development.xlsx',sheet_name='FL')
data=data.set_index('Tool')
table_data=data.T
chart_data=data

# No Percent data
np_data=table_data[['TN Victor2 Pro White','TN Victor2 Pro Silver']]
npp_data=table_data[['VH Victor2 Pro White','VH Victor2 Pro Silver']]
np_data=np_data.round(1)
npp_data=npp_data.round(1)

# Percent data
p_data=table_data[['Victor2 Pro White Diff Rate(%)']]
pp_data=table_data[['Victor2 Pro Silver Diff Rate(%)']]
p_data=p_data.round(2)
pp_data=pp_data.round(2)

# 데이터 넣기
axF1=np_data.plot(use_index=True, color=['#D2D6D1','#D2D6D1'],
                  marker='o',markersize=4,
                  label=['TN Victor2 Pro White','TN Victor2 Pro Silver'],
                  ax=ax[1],legend=None)
axF1=npp_data.plot(use_index=True, color=['#EE2113','#EE2113'],
                  marker='o',markersize=4,
                  label=['VH Victor2 Pro White','VH Victor2 Pro Silver'],
                  ax=ax[1],legend=None) 

axF2 = axF1.twinx()

p_data.plot(kind='line', use_index=True,legend=None,
            color='#7CF9A2',
            marker='o',markersize=4,
            label=['Victor2 Pro White Diff Rate(%)'],ax=axF2)

pp_data.plot(kind='line', use_index=True,legend=None,
            color='#B6F97C',
             marker='o',markersize=4,
            label=['Victor2 Pro Silver Diff Rate(%)'],ax=axF2)




# UI
x = np.arange(18)
y1=list(np.array(table_data['TN Victor2 Pro White'].round(1).tolist()))
y2=list(np.array(table_data['VH Victor2 Pro White'].round(1).tolist()))
y3=list(np.array(table_data['Victor2 Pro White Diff Amount'].round(1).tolist()))
y4=list(np.array((table_data['Victor2 Pro White Diff Rate(%)']).round(2).tolist()))
y5=list(np.array(table_data['TN Victor2 Pro Silver'].round(1).tolist()))
y6=list(np.array(table_data['VH Victor2 Pro Silver'].round(1).tolist()))
y7=list(np.array(table_data['Victor2 Pro Silver Diff Amount'].round(1).tolist()))
y8=list(np.array((table_data['Victor2 Pro Silver Diff Rate(%)']).round(2).tolist()))


a =[23.225 ,23.188, 23.173, 23.183, 23.186, 23.174, 23.129, 23.078, 23.025, 23.065, 23.068, 23.061, 22.998, 23.007 , 22.849, 22.765, 22.749]
b='kVND/$'
#
for i in range(18):
    axF1.annotate(y1[i],xy=(x[i],y1[i]-0.02),va='top',color='#B5A9A7',fontsize=9)
    axF1.annotate(y2[i],xy=(x[i],y2[i]-0.02),va='top',color='red',fontsize=9)
    axF2.annotate(str(y4[i])+"%",xy=(x[i],y4[i]-0.02),va='top',color='#00B050',fontsize=9)
    axF1.annotate(y5[i],xy=(x[i],y5[i]-0.02),va='bottom',color='#B5A9A7',fontsize=9)
    axF1.annotate(y6[i],xy=(x[i],y6[i]-0.02),va='top',color='red',fontsize=9)
    axF2.annotate(str(y8[i])+"%",xy=(x[i],y8[i]-0.02),va='bottom',color='#92D050',fontsize=9)
    plt.text(x[i]-0.4, 42 ,a[i],color='green')
    i=i+1

plt.text(-1.5, 42 , b,color='green')
    

plt.title("Material Comparsion with VH(Sep 2W)",fontsize=13, pad='30')
axF1.set_ylabel('USD($)',color='gray')
axF2.set_ylabel('Percentage(%)',color='gray')
axF2.yaxis.set_major_formatter(mtick.PercentFormatter()) # y축 퍼센트로 만들기
axF1.set_ylim(200,320)
axF2.set_ylim(-5,40)
axF2.set_xticks(np.arange(18))
axF2.set_xticklabels(list(p_data.index))
axF2.set_xlim([-0.5,17.5])
axF1.legend(loc='upper left')
axF2.legend(loc='upper right',bbox_to_anchor=(0.41,1))

# 차트 만들기
ax[1].set_axis_off()

array1=np.full(18,'#F8E0F6') # col 색깔 채우기
array2=['#D2D6D1','#F6CCC8','white','#A3E7B3','#D2D6D1','#F6CCC8','white','#D8FDC9'] # row 색깔 채우기

the_table = plt.table(cellText=chart_data.round(2).values,
                      rowLabels=chart_data.index,
                      colLabels=chart_data.columns,
                      colColours=array1,
                      rowColours=array2)


plt.subplots_adjust(left=0.17,bottom=-0.1)
the_table.auto_set_font_size(False)
the_table.set_fontsize(9)


plt.savefig('FL.png')

#Body 꾸미기
text0='This is DX activities from LGEUS R&D Team\nPerson in charge: LGEUS R&D Team Eunbi Yoon\n\n\n'
msg.attach(MIMEText(text0,'plain'))

#첨부 파일1
with open('FL.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('FL_1.png'))
msg.attach(image)



#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")
