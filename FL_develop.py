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
########################## 전체 틀 만들기
#그래프 만들기
fig = plt.figure(constrained_layout=True,figsize=(18,13))
gs = fig.add_gridspec(16,16)
ax0 = fig.add_subplot(gs[0:3,1:9])
ax1 = fig.add_subplot(gs[1:2,12:16])
ax2 = fig.add_subplot(gs[3:4,12:16])
ax3 = fig.add_subplot(gs[4:9,0:16])


######################## ax[0]
# Chart data 읽기
data0=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='TL_0')
data0.index=data0['Changing reason']
data0=data0.drop(['Changing reason'],axis=1)
ax0.set_axis_off()

col_Colours=np.full(11,'#CCE7E7')
row_Colours=['#F6CCC8','#F6CCC8','#F6CCC8','#D6D6D6']
table0 = ax0.table(cellText=data0.values,
                      rowLabels=data0.index,colLabels=data0.columns,colColours=col_Colours, rowColours=row_Colours,loc='center')

table0.auto_set_font_size(False)
table0.set_fontsize(9)
table0.auto_set_column_width(col=list(range(len(data0.columns))))

ax0.set_title("Price Change History ($)",fontsize=11,x=0.05,y=0.65) ## title
plt.text(0,5, 'TH Increased $6.66 (Exchange Rate $6.3↑ Material Cost Increase $0.33↑), TN increased $1.96 (Material Increase $1.75↑)',fontsize=11,color='green') ## title



######################## ax[1]
data1=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='TL_1')
data1.index=data1['Increasing Month']
data1=data1.drop('Increasing Month',axis=1)
ax1.set_axis_off()


col_Colours=np.full(2,'#CCE7E7')
row_Colours=['#F6CCC8','#F6CCC8','#F6CCC8','#F6CCC8','#B6ACCE']

table1 = ax1.table(cellText=data1.values,
                      rowLabels=data1.index,colLabels=data1.columns,loc='top',colColours=col_Colours, rowColours=row_Colours)

ax1.set_title('VH Site',x=0.05,y=1.3,color='green',fontsize=11)

table1.auto_set_font_size(False)
table1.set_fontsize(9)
table1.auto_set_column_width(col=list(range(len(data1.columns))))



######################## ax[2]
data2=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='TL_2')
data2.index=data2['Increasing Month']
data2=data2.drop('Increasing Month',axis=1)
ax2.set_axis_off()

col_Colours=np.full(2,'#CCE7E7')
row_Colours=['#D6D6D6','#D6D6D6','#D6D6D6','#B6ACCE']
table2 = ax2.table(cellText=data2.values,
                      rowLabels=data2.index,colLabels=data2.columns,loc='top',colColours=col_Colours, rowColours=row_Colours)

ax2.set_title('TN Site',x=0.05,y=1.9,color='green',fontsize=11)

table2.auto_set_font_size(False)
table2.set_fontsize(9)
table2.auto_set_column_width(col=list(range(len(data2.columns))))



######################## ax[3]
# data읽기
data=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='TL')
data=data.set_index('Tool')
table_data=data.T
chart_data=data

# No Percent data
np_data=table_data[['TD Pro White TN','TD Pro Silver TN']]
npp_data=table_data[['TD Pro White TH','TD Pro Silver TH']]
np_data=np_data.round(1)
npp_data=npp_data.round(1)

# Percent data
p_data=table_data[['TD Pro Silver Diff Rate(%)']]
pp_data=table_data[['TD Pro White Diff Rate(30.8)(%)','TD Pro White Diff Rate(%)']]
p_data=p_data.round(2)
pp_data=pp_data.round(2)


# 데이터 넣기
axF1=np_data.plot(use_index=True, color=['#D2D6D1','#D2D6D1'],
                  marker='o',markersize=4,
                  label=['TD Pro White TN','TD Pro Silver TN'],ax=ax3,legend=None)

axF1=npp_data.plot(use_index=True, color=['#F6CCC8','#F6CCC8'],
                  marker='o',markersize=4,
                  label=['TD Pro White TH','TD Pro Silver TH'],ax=ax3,legend=None)

axF2 = axF1.twinx()

p_data.plot(kind='line', use_index=True,legend=None,
            color='#C8F6A0',
            marker='o',markersize=4,
            label=['TD Pro Silver Diff Rate(%)'],ax=axF2)

pp_data.plot(kind='line', use_index=True,legend=None,
            color='#C4F3D2',linestyle='--',
             marker='o',markersize=4,
            label=['TD Pro White Diff Rate(30.8)(%)'],ax=axF2)

pp_data.plot(kind='line', use_index=True,legend=None,
            color='#C4F3D2',
             marker='o',markersize=4,
            label=['TD Pro White Diff Rate(%)'],ax=axF2)



# UI
x = np.arange(23)
y1=list(np.array(table_data['TD Pro White TN'].round(1).tolist()))
y2=list(np.array(table_data['TD Pro White TH'].round(1).tolist()))
y3=list(np.array(table_data['TD Pro Silver TN'].round(1).tolist()))
y4=list(np.array((table_data['TD Pro Silver TH']).round(2).tolist()))
y5=list(np.array(table_data['TD Pro White Diff Rate(%)'].round(1).tolist()))
y6=list(np.array(table_data['TD Pro White Diff Rate(30.8)(%)'].round(1).tolist()))
y7=list(np.array(table_data['TD Pro Silver Diff Rate(%)'].round(1).tolist()))


a =[30.3,31.3,32.0,32.7,32.1,32.2,31.4,31.3,31.4,31.3,30.5,30.1,30.0,30.0,30.7,31.3,31.3,31.4,32.6,33.5,33.3,33.3,33.3]
b='kVND/$'
#
for i in range(23):
    axF1.annotate(y1[i],xy=(x[i],y1[i]),va='top',color='gray',fontsize=9)
    axF1.annotate(y2[i],xy=(x[i],y2[i]-0.02),va='bottom',color='red',fontsize=9)
    axF1.annotate(y3[i],xy=(x[i],y3[i]),va='top',color='gray',fontsize=9)
    axF1.annotate(y4[i],xy=(x[i],y4[i]-0.02),va='bottom',color='red',fontsize=9)
    axF2.annotate(str(y5[i])+"%",xy=(x[i],y5[i]-0.02),va='bottom',color='#00B050',fontsize=9)
    axF2.annotate(str(y6[i])+"%",xy=(x[i],y6[i]),va='bottom',color='#629E2E',fontsize=9)
    axF2.annotate(str(y7[i])+"%",xy=(x[i],y7[i]-0.02),va='bottom',color='#629E2E',fontsize=9)

    plt.text(x[i]-0.4, 72 ,a[i],color='green')
    i=i+1

plt.text(-1.5, 72 , b,color='green')
    

plt.title("Material comparison with TH (Nov 1W)",fontsize=13, pad='25')
axF1.set_ylabel('USD($)',color='gray')
axF2.set_ylabel('Percentage(%)',color='gray')
axF2.yaxis.set_major_formatter(mtick.PercentFormatter()) # y축 퍼센트로 만들기

axF1.set_ylim(90,290)
axF2.set_ylim(-10,70)
axF2.set_xticks(np.arange(23))

axF2.set_xticklabels(list(p_data.index))
axF2.set_xlim([-0.5,22.5])
axF1.legend(loc='upper left')
axF2.legend(loc='upper right',bbox_to_anchor=(0.41,1))

# 차트 만들기
ax3.set_axis_off()

array1=np.full(23,'#CCE7E7') # col 색깔 채우기
array2=['#D2D6D1','#F6CCC8','#F6CCC8','#A3E7B3','#A3E7B3','#A3E7B3','#A3E7B3','white','#F6CCC8','#F6CCC8','#D8FDC9','#D8FDC9'] # row 색깔 채우기

the_table = plt.table(cellText=chart_data.round(2).values,
                      rowLabels=chart_data.index,
                      colLabels=chart_data.columns,
                      colColours=array1,
                      rowColours=array2)

plt.subplots_adjust(left=0.17,bottom=-0.1)
the_table.auto_set_font_size(False)
the_table.set_fontsize(9)
axF1.get_legend().remove()
axF2.get_legend().remove()
#plt.show()
plt.savefig('TL.png')


########################## FL 만들기 ################################
########################## 전체 틀 만들기
#그래프 만들기
fig = plt.figure(constrained_layout=True,figsize=(16,9))
gs = fig.add_gridspec(14,16)
ax0 = fig.add_subplot(gs[0:3,3:11])
ax1 = fig.add_subplot(gs[1:2,12:16])
ax2 = fig.add_subplot(gs[3:4,12:16])
ax3 = fig.add_subplot(gs[4:9,0:16])


######################## ax[0]
# Chart data 읽기
data0=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='FL_0')
data0.index=data0['Changing reason']
data0=data0.drop(['Changing reason'],axis=1)
ax0.set_axis_off()

col_Colours=np.full(11,'#CCE7E7')
row_Colours=['#F6CCC8','#F6CCC8','#F6CCC8','#D6D6D6']
table0 = ax0.table(cellText=data0.values,
                      rowLabels=data0.index,colLabels=data0.columns,colColours=col_Colours, rowColours=row_Colours,loc='center')

table0.auto_set_font_size(False)
table0.set_fontsize(9)
table0.auto_set_column_width(col=list(range(len(data0.columns))))

ax0.set_title("Price Change History ($)",fontsize=11,x=-0.2,y=0.7) ## title
plt.text(0,0,'VH : $0.01↑ (Exchange Rate: $0.01↑ Material Cost Increase $0↑), TN : $0.43↓ (Material : $0.43↓)',fontsize=11,color='green') ## title



######################## ax[1]
data1=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='FL_1')
data1.index=data1['Increasing Month']
data1=data1.drop('Increasing Month',axis=1)
ax1.set_axis_off()


col_Colours=np.full(2,'#CCE7E7')
row_Colours=['#F6CCC8','#F6CCC8','#F6CCC8','#F6CCC8','#B6ACCE']

table1 = ax1.table(cellText=data1.values,
                      rowLabels=data1.index,colLabels=data1.columns,loc='top',colColours=col_Colours, rowColours=row_Colours)

ax1.set_title('VH Site',x=-0.1,y=2.6,color='green',fontsize=11)

table1.auto_set_font_size(False)
table1.set_fontsize(9)
table1.auto_set_column_width(col=list(range(len(data1.columns))))



######################## ax[2]
data2=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='FL_2')
data2.index=data2['Increasing Month']
data2=data2.drop('Increasing Month',axis=1)
ax2.set_axis_off()

col_Colours=np.full(2,'#CCE7E7')
row_Colours=['#D6D6D6','#D6D6D6','#D6D6D6']
table2 = ax2.table(cellText=data2.values,
                      rowLabels=data2.index,colLabels=data2.columns,loc='top',colColours=col_Colours, rowColours=row_Colours)

ax2.set_title('TN Site',x=-0.1,y=2.2,color='green',fontsize=11)

table2.auto_set_font_size(False)
table2.set_fontsize(9)
table2.auto_set_column_width(col=list(range(len(data2.columns))))



######################## ax[3]
# data읽기
data=pd.read_excel('C:/Users/eunbi1.yoon/AppData/Local/Programs/Python/Python37/Request/Development.xlsx',sheet_name='FL')
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
                  ax=ax3,legend=None)
axF1=npp_data.plot(use_index=True, color=['#F6CCC8','#F6CCC8'],
                  marker='o',markersize=4,
                  label=['VH Victor2 Pro White','VH Victor2 Pro Silver'],
                  ax=ax3,legend=None) 

axF2 = axF1.twinx()

p_data.plot(kind='line', use_index=True,legend=None,
            color='#C4F3D2',
            marker='o',markersize=4,
            label=['Victor2 Pro White Diff Rate(%)'],ax=axF2)

pp_data.plot(kind='line', use_index=True,legend=None,
            color='#C8F6A0',
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


a =[23.225 ,23.188, 23.173, 23.183, 23.186, 23.174, 23.129, 23.078, 23.025, 23.065, 23.068, 23.061, 22.998, 23.007 , 22.849, 22.765, 22.755, 22.755]
b='kVND/$'
#
for i in range(18):
    axF1.annotate(y1[i],xy=(x[i],y1[i]-0.02),va='top',color='gray',fontsize=9)
    axF1.annotate(y2[i],xy=(x[i],y2[i]-0.02),va='bottom',color='red',fontsize=9)
    axF2.annotate(str(y4[i])+"%",xy=(x[i],y4[i]-0.02),va='top',color='#00B050',fontsize=9)
    axF1.annotate(y5[i],xy=(x[i],y5[i]-0.02),va='bottom',color='gray',fontsize=9)
    axF1.annotate(y6[i],xy=(x[i],y6[i]-0.02),va='bottom',color='red',fontsize=9)
    axF2.annotate(str(y8[i])+"%",xy=(x[i],y8[i]-0.02),va='bottom',color='#629E2E',fontsize=9)
    plt.text(x[i]-0.4, 52 ,a[i],color='green')
    i=i+1

plt.text(-1.5, 52 , b,color='green')
    

plt.title("Material comparison with VH (Nov 1W)",fontsize=13, pad='25')
axF1.set_ylabel('USD($)',color='gray')
axF2.set_ylabel('Percentage(%)',color='gray')
axF2.yaxis.set_major_formatter(mtick.PercentFormatter()) # y축 퍼센트로 만들기
axF1.set_ylim(180,340)
axF2.set_ylim(-10,50)
axF2.set_xticks(np.arange(18))

axF2.set_xticklabels(list(p_data.index))
axF2.set_xlim([-0.5,17.5])
axF1.legend(loc='upper left')
axF2.legend(loc='upper right',bbox_to_anchor=(0.41,1))

# 차트 만들기
ax3.set_axis_off()

array1=np.full(18,'#CCE7E7') # col 색깔 채우기
array2=['#D2D6D1','#F6CCC8','white','#A3E7B3','#D2D6D1','#F6CCC8','white','#D8FDC9'] # row 색깔 채우기

the_table = plt.table(cellText=chart_data.round(2).values,
                      rowLabels=chart_data.index,
                      colLabels=chart_data.columns,
                      colColours=array1,
                      rowColours=array2)

plt.subplots_adjust(left=0.17,bottom=-0.1)
the_table.auto_set_font_size(False)
the_table.set_fontsize(9)

#plt.show()
plt.savefig('FL.png')

#Body 꾸미기
text0='This is DX activities from LGEUS R&D Team\nPerson in charge: LGEUS R&D Team Eunbi Yoon\n\n\n'
msg.attach(MIMEText(text0,'plain'))

#첨부 파일1
with open('TL.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('TL.png'))
msg.attach(image)

#첨부 파일1
with open('FL.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('FL.png'))
msg.attach(image)



#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")
