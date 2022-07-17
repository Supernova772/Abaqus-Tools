#!/user/bin/python
#-*-coding:UTF-8-*-
from abaqus import *
from abaqusConstants import *
from caeModules import *
from driverUtils import executeOnCaeStartup
import numpy as np
from visualization import *
from odbAccess import *
import xlwt
import math
import xlwings as xw

import threading
import time
odbname,wbname,SH,ns1name,ns2name,ns3name,ns4name,ns5name= getInputs(
  fields=(('odbname:', '3sidedub356x171x45-d380.odb'),('wbname:', '356x171x45d380'),('Heading:', 'Temperature'),('Nodeset 1:', 'UFLANGE'),('Nodeset 2:','BFLANGE'),('Nodeset 3:','TEMP'),('Nodeset 4:','HFU'),('Nodeset 5:','HFB')),
  label='Enter information',
  dialogTitle='Average temperature of node sets')
ns1name=str(ns1name)
ns2name=str(ns2name)
o3 = session.openOdb(name='c:/temp/'+odbname)
session.viewports['Viewport: 1'].setValues(displayedObject=o3)
session.viewports['Viewport: 1'].makeCurrent()
odb=openOdb(path='c:/temp/'+odbname)
wbkName=wbname
wbk=xlwt.Workbook()
sheet=wbk.add_sheet('Sheet1')

frameRepository=odb.steps['Step-1'].frames

iframes=len(frameRepository)
RefDataTime=np.zeros((iframes,1))
RefDataTemp=np.zeros((iframes,1))
RefDataHFL=np.zeros((iframes,1))
ListDataTime=np.zeros((iframes,1))
top=np.zeros((1,2))
bottom=np.zeros((1,2))
precision=2
precision1=6
sheet.write(0,0,(SH))
for j in range(3):
  if j==1:
    ns1name=ns2name
  elif j==2:
    ns1name=ns3name
  elif j==3:
    ns1name=ns4name
  elif j==4:
    ns1name=ns5name

  if j<3 :
      sheet.write(2,2*j,('Time/s'))
      sheet.write(2,2*j+1,('Temperature/C'))
      sheet.write(1,2*j,(ns1name))
  elif j>=3:
      sheet.write(2,2*j,('Time/s'))
      sheet.write(2,2*j+1,('HeatFlux'))
      sheet.write(1,2*j,(ns1name))
 
 
  RefPointSet=odb.rootAssembly.nodeSets[ns1name]
  
  
  
  #xyList=session.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=(('NT11', NODAL), ),nodeSets=(ns2name,))
  #xyList=session.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=(('NT11', NODAL), ),nodeSets=(ns3name,))
  if j<3 :
    xyList=session.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=(('NT11', NODAL), ),nodeSets=(ns1name,))
    for i in range(iframes):
        Temp=frameRepository[i].fieldOutputs['NT11']
        RefTemp=Temp.getSubset(region=RefPointSet)
        RefTempValue=RefTemp.values 
        num=len(RefTempValue)
        overall=0
        Temp=np.zeros((num,1))
        for k in range(num):
         middle=RefTempValue[k]
         Temp[k]=middle.data
         overall=overall+Temp[k]
         #a=RefTempA.data
         #b=RefTempB.data
        RefDataTemp[i]=overall/num
        #sheet.write(i+1,0,round(RefDataTime[i],precision)
        sheet.write(i+3,2*j+1,round(RefDataTemp[i],precision))
    
        #a1=xyList[0].data
        #a2=xyList[1].data
        #top=a1[i]
        #bottom=a2[i]
        ListDataTime[i]=xyList[0][i][0]
        sheet.write(i+3,2*j,round(ListDataTime[i],precision))
  
       
   
  elif j>=3 :
    xyList=session.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=((
    'HFL', INTEGRATION_POINT, ((COMPONENT, 'HFL2'), )), ),nodeSets=(ns1name,))
    for i in range(iframes):
        #HFL=frameRepository[i].fieldOutputs['HFL']
        #RefHFL=HFL.getSubset(region=RefPointSet)
        #RefHFLValue=RefHFL.values 
        num=len(xyList)
        overall=0
        HFL=np.zeros((num,1))
        for k in range(num):
             middle=xyList[k][i][1]
             HFL[k]=middle
             overall=overall+HFL[k]
             #a=RefHFLA.data
             #b=RefHFLB.data
        RefDataHFL[i]=overall
        #sheet.write(i+1,0,round(RefDataTime[i],precision)
        sheet.write(i+3,2*j+1,round(RefDataHFL[i],precision1))
        #a1=xyList[0].data
        #a2=xyList[1].data
        #top=a1[i]
        #bottom=a2[i]
        ListDataTime[i]=xyList[0][i][0]
        sheet.write(i+3,2*j,round(ListDataTime[i],precision))  
  sheet.write(iframes+3,2*j,num)

wbk.save(wbkName+'.xls')

#import win32api
#win32api.ShellExecute(0, 'open', 'D:\\temp\\'+wbkName+'.xls', '','',1)
a=['','']
dir1=['']
format1=['']
app=xw.App(visible=True,add_book=False)
#open workbook
dir1[0]="c:\\temp\\"
format1[0]=".xls"
##databook file
a[0]=dir1[0]+wbkName+format1[0]
##workbook file
#a[1]='C:\Users\y24884yl\OneDrive - The University of Manchester\Research\3sided temp based on EC4 Temperature and verification.xlsm '
def open_book(i):
  app.books.open(a[i])
#open workbook
time.sleep(1)
i=0
open_book(i)


