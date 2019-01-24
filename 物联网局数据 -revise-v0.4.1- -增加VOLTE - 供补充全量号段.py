#! /usr/bin/env python 
# -*- coding:utf-8 –*-  

#Gx/SLh接口先不做！！！！！


import xlrd
import xlwt
import os
workpath = os.getcwd()
import datetime
nowTime=datetime.datetime.now().strftime('%H-%M-%S')
def run(rootPath):


	data_dir= os.path.abspath(rootPath)+"\\待制作文件\\"
	#for file1 in os.listdir(rootPath):
	#	print (file1)
	for root,dirs,files in os.walk(data_dir):#遍历文件夹
		for file_node in files:
			file_Name=os.path.join(root,file_node)
			if(str(file_Name).endswith("xls") or str(file_Name).endswith("xlsx"))\
				and (str(file_node).find("物联网启用LTE用户号码批复表")!=-1):#物联网启用LTE用户号码批复表——IMSI/
				VOLTE_WLW_IMSI1(file_Name)
				#VOLTE_WLW_MSISDN(file_Name,jsj_name)
			if(str(file_Name).endswith("xls") or str(file_Name).endswith("xlsx"))\
				and (str(file_node).find("中国移动网内启用物联网号段批复表")!=-1):#物联网启用LTE用户号码批复表——IMSI/
				VOLTE_WLW_MSISDN1(file_Name)

def VOLTE_WLW_IMSI(file_Name,jsj_name):
	Province=["安徽","北京","福建","甘肃","广东","广西","贵州","海南","河北","河南","黑龙江","湖北",	"湖南","吉林","江苏","江西","辽宁","内蒙古","宁夏","青海","山东","山西","陕西","上海","四川","天津","西藏","新疆","云南","浙江","重庆"]
	#BJ=NJ=JN=SJ=CD=GZ=HZ=WH=ZZ=HW=""
	HW=""
	#HDRA_Prov=["BJ","NJ","JN","SJ","CD","GZ","HZ","WH","ZZ"]
	HDRA_Prov_CH=["北京","河北","浙江","江苏","四川","河南","湖北","广东","山东"]
	#HDRA_Prov_WLW={"北京","浙江","四川","广东"}
	'''are信息已经没用了'''
	北京={'name_en':"BEIJING",'name_short':"BJ",'rt':"",'hss':["BFMHSS01AZX"],'area':["北京","天津","河北","山西","内蒙古","辽宁","吉林","黑龙江","甘肃","山东","政企"]}
	河北={'name_en':"HEBEI",'name_short':"SJ",'rt':"",'hss':[],'area':[]}
	河南={'name_en':"HENAN",'name_short':"ZZ",'rt':"",'hss':[],'area':[]}
	江苏={'name_en':"JIANGSU",'name_short':"NJ",'rt':"",'hss':[],'area':[]}
	山东={'name_en':"SHANDONG",'name_short':"JN",'rt':"",'hss':[],'area':[]}
	浙江={'name_en':"ZHEJIANG",'name_short':"HZ",'rt':"",'hss':["DFMHSS01FE01AZX","DFMHSS02FE01AZX","DFMHSS03FE01AZX","DFMHSS04FE01AZX","DFMHSS05FE01AZX"],'area':["上海","浙江","江苏"]}
	四川={'name_en':"SICHUAN",'name_short':"CD",'rt':"",'hss':["XFMHSS01FE01AHW","XFMHSS02FE01AHW","XFMHSS03FE01AHW"],'area':["重庆","四川","陕西","云南","西藏","新疆","河南","湖北","青海","宁夏","物联网"]}
	广东={'name_en':"GUANGDONG",'name_short':"GZ",'rt':"",'hss':["NFMHSS01AHW"],'area':["广东","广西","海南","贵州","湖南","安徽","江西","福建"]}
	湖北={'name_en':"HUBEI",'name_short':"WH",'rt':"",'hss':[],'area':[]}

	workbook = xlrd.open_workbook(file_Name)
	table=workbook.sheets()[0]
	for i in range(1,table.nrows):
		print ("正在制作第"+str(i)+"条"+",共"+str(table.nrows-1)+"条")
		prov=table.cell(i,0).value
		imsi_value=str(int(table.cell(i,1).value))
		hss_value=str(table.cell(i,2).value)
		for item in HDRA_Prov_CH:
			if  hss_value in eval(item)['hss']:
				eval(item)['rt']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETDIR=ROUTESET|{}-rs.\n".format(imsi_value,hss_value)
				route=ROUTE(prov,item,'BELL','')
				route_HW=ROUTE(prov,eval(item)['name_en'],'HW',table.cell(i,2).value)
				HW+="ADD RTIMSI: REFERINDEX=0,IMSI=\""+imsi_value+"\",NEXTRULE=RTEXIT,NEXTINDEX="+str(route_HW)+",MOG=\""+route+"\";{"+eval(item)['name_short']+"DRA01AHW-A}\n"
			else:
				route=ROUTE(prov,item,'BELL','')
				eval(item)['rt']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETDIR=ROUTESET|{}-rs.\n".format(imsi_value,route)
				route_HW=ROUTE(prov,eval(item)['name_en'],'HW','')
				HW+="ADD RTIMSI: REFERINDEX=0,IMSI=\""+imsi_value+"\",NEXTRULE=RTEXIT,NEXTINDEX="+str(route_HW)+",MOG=\""+route+"\";{"+eval(item)['name_short']+"DRA01AHW-A}\n"
	for item in HDRA_Prov_CH:
		c=open(workpath+"\\脚本\\"+"["+eval(item)['name_short']+"DRA01AAL-B]B_"+jsj_name+".txt",'w')
		c.write(eval(item)['rt'])
		c.close()
	c=open(workpath+"\\脚本\\"+"_[HW_HDRA]H_"+jsj_name+".txt",'w')
	c.write(HW)
	c.close()

def VOLTE_WLW_IMSI1(file_Name):
	route_bell,route_hw=ROUTE_NEW()
	workbook = xlrd.open_workbook(file_Name)
	table=workbook.sheets()[0]
	HW=""
	HDRA_Prov_CH={"北京":{'name_en':"BEIJING",'name_short':"BJ",'jsj':"",'hss':["BFMHSS01AZX"],'area':["北京","天津","河北","山西","内蒙古","辽宁","吉林","黑龙江","甘肃","山东","政企"]},\
					"河北":{'name_en':"HEBEI",'name_short':"SJ",'jsj':"",'hss':[],'area':[]},\
					"河南":{'name_en':"HENAN",'name_short':"ZZ",'jsj':"",'hss':[],'area':[]},\
					"江苏":{'name_en':"JIANGSU",'name_short':"NJ",'jsj':"",'hss':[],'area':[]},\
					"山东":{'name_en':"SHANDONG",'name_short':"JN",'jsj':"",'hss':[],'area':[]},\
					"浙江":{'name_en':"ZHEJIANG",'name_short':"HZ",'jsj':"",'hss':["DFMHSS01FE01AZX","DFMHSS02FE01AZX","DFMHSS03FE01AZX","DFMHSS04FE01AZX","DFMHSS05FE01AZX"],'area':["上海","浙江","江苏"]},\
					"四川":{'name_en':"SICHUAN",'name_short':"CD",'jsj':"",'hss':["XFMHSS01FE01AHW","XFMHSS02FE01AHW","XFMHSS03FE01AHW"],'area':["重庆","四川","陕西","云南","西藏","新疆","河南","湖北","青海","宁夏","物联网"]},\
					"广东":{'name_en':"GUANGDONG",'name_short':"GZ",'jsj':"",'hss':["NFMHSS01AHW"],'area':["广东","广西","海南","贵州","湖南","安徽","江西","福建"]},\
					"湖北":{'name_en':"HUBEI",'name_short':"WH",'jsj':"",'hss':[],'area':[]}}
	for i in range(1,table.nrows):
		print ("正在制作第"+str(i)+"条"+",共"+str(table.nrows-1)+"条")
		prov=table.cell(i,0).value
		imsi_value=str(int(table.cell(i,1).value))
		hss_value=str(table.cell(i,2).value)
		
		for item in HDRA_Prov_CH.keys():
			#if hss_value in HDRA_Prov_CH[item]["hss"]:	#物联网四大区HDRA制作内容
			if prov in HDRA_Prov_CH[item]["area"]:   #####由于有现网号段的HSS信息不准，恢复按照省份判断
				#BELL
				#S6a指向EPC HSS，现保留原始格式 / Cx Zh指向IMS的HSS   2019.01.22 存量号段不增加S6a接口
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETAPP=S6a|ROUTESET|{}-rs.\n".format(imsi_value,hss_value)
				HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETAPP=Cx|ROUTESET|{}-IMS-rs.\n".format(imsi_value,hss_value)
				HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETAPP=Zh|ROUTESET|{}-IMS-rs.\n".format(imsi_value,hss_value)
				#print (HDRA_Prov_CH[item]['jsj'])
				#华为
				#Cx生成一条IPMU(IMSI) Zh生成一条IMPI（IMSI）S6a生成IMSI
				#route=ROUTE(prov,item,'BELL','')
				#route_HW=ROUTE(prov,eval(item)['name_en'],'HW',table.cell(i,2).value)
				#print (route_hw[HDRA_Prov_CH[item]['name_en']])
				nextindex=route_hw[HDRA_Prov_CH[item]['name_en']][hss_value]  #指向IMS-HSS
				nextindex_ims=route_hw[HDRA_Prov_CH[item]['name_en']][hss_value+'-IMS']
				mog=hss_value
				#mog_ims=hss_value+'-IMS'  2019.01.22更新 不加IMS了
				mog_ims=hss_value
				#S6a接口  EPC-HSS    2019.01.22更新——存量号段不增加IMSI
				#HW+="ADD RTIMSI: REFERINDEX=0,IMSI=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				#Cx接口 IMS-HSS  ！！！
				HW+="ADD RTIMPU: REFERINDEX=0,IMPU=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#Zh接口 IMS-HSS  ！！！
				HW+="ADD RTIMPI: REFERINDEX=0,IMPI=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
			else:
				#BELL
				#S6a指向EPC HSS，现保留原始格式 / Cx Zh指向IMS的HSS
				rt_bl=route_bell[prov]["归属物联网大区"]
				HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETDIR=ROUTESET|{}-rs.\n".format(imsi_value,rt_bl)


				#非落地省制作一条？？？？？？？？？？？？？
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETAPP=S6a|ROUTESET|{}-rs.\n".format(imsi_value,rt_bl)
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETAPP=Cx|ROUTESET|{}-IMS-rs.\n".format(imsi_value,rt_bl)
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:IMSI={},TARGETAPP=Zh|ROUTESET|{}-IMS-rs.\n".format(imsi_value,rt_bl)
				#print (HDRA_Prov_CH[item]['jsj'])
				#华为
				#Cx生成一条IPMU(IMSI) Zh生成一条IMPI（IMSI）S6a生成IMSI
				#route=ROUTE(prov,item,'BELL','')
				#route_HW=ROUTE(prov,eval(item)['name_en'],'HW',table.cell(i,2).value)
				#print (route_hw[HDRA_Prov_CH[item]['name_en']])
				prov_belong=route_bell[prov]["归属物联网大区"]
				#print (prov_belong)
				nextindex=route_hw[HDRA_Prov_CH[item]['name_en']][prov_belong]
				mog=hss_value
				#？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？这边好像要指到EPC-HSS
				#S6a接口  EPC-HSS   2019.01.22  存量号段不生成S6a
				#HW+="ADD RTIMSI: REFERINDEX=0,IMSI=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				#Cx接口 IMS-HSS  ！！！  此处不涉及，指向归属大区
				HW+="ADD RTIMPU: REFERINDEX=0,IMPU=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				#Zh接口 IMS-HSS  ！！！  此处不涉及，指向归属大区
				HW+="ADD RTIMPI: REFERINDEX=0,IMPI=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])

	for item in HDRA_Prov_CH.keys():
		c=open(workpath+"\\脚本\\"+"["+HDRA_Prov_CH[item]['name_short']+"DRA01AAL-B]B_"+nowTime+".txt",'w')
		c.write(HDRA_Prov_CH[item]['jsj'])
		c.close()
	c=open(workpath+"\\脚本\\"+"_[HW_HDRA]H_"+nowTime+".txt",'w')
	c.write(HW)
	c.close()


def VOLTE_WLW_MSISDN1(file_Name):
	route_bell,route_hw=ROUTE_NEW()
	workbook = xlrd.open_workbook(file_Name)
	table=workbook.sheets()[0]
	HW=""
	HDRA_PROV_WLW=["北京","浙江","广东","四川"]
	HDRA_Prov_CH={"北京":{'name_en':"BEIJING",'name_short':"BJ",'jsj':"",'hss':["BFMHSS01AZX"],'area':["北京","天津","河北","山西","内蒙古","辽宁","吉林","黑龙江","甘肃","山东","政企"],'pcrf':'BFMPCRF0102AZX'},\
					"河北":{'name_en':"HEBEI",'name_short':"SJ",'jsj':"",'hss':[],'area':[]},\
					"河南":{'name_en':"HENAN",'name_short':"ZZ",'jsj':"",'hss':[],'area':[]},\
					"江苏":{'name_en':"JIANGSU",'name_short':"NJ",'jsj':"",'hss':[],'area':[]},\
					"山东":{'name_en':"SHANDONG",'name_short':"JN",'jsj':"",'hss':[],'area':[]},\
					"浙江":{'name_en':"ZHEJIANG",'name_short':"HZ",'jsj':"",'hss':["DFMHSS01FE01AZX","DFMHSS02FE01AZX","DFMHSS03FE01AZX","DFMHSS04FE01AZX","DFMHSS05FE01AZX"],'area':["上海","浙江","江苏"],'pcrf':'DFMPCRFPOOLER'},\
					"四川":{'name_en':"SICHUAN",'name_short':"CD",'jsj':"",'hss':["XFMHSS01FE01AHW","XFMHSS02FE01AHW","XFMHSS03FE01AHW"],'area':["重庆","四川","陕西","云南","西藏","新疆","河南","湖北","青海","宁夏","物联网"],'pcrf':'XFMPCRFPOOLHW'},\
					"广东":{'name_en':"GUANGDONG",'name_short':"GZ",'jsj':"",'hss':["NFMHSS01AHW"],'area':["广东","广西","海南","贵州","湖南","安徽","江西","福建"],'pcrf':'NFMPCRF0102AZX'},\
					"湖北":{'name_en':"HUBEI",'name_short':"WH",'jsj':"",'hss':[],'area':[]}}
	for i in range(1,table.nrows):
		print ("正在制作第"+str(i)+"条"+",共"+str(table.nrows-1)+"条")
		prov=table.cell(i,0).value
		imsi_value=str(int(table.cell(i,1).value))
		hss_value=str(table.cell(i,2).value)
		
		for item in HDRA_Prov_CH.keys():
			#if hss_value in HDRA_Prov_CH[item]["hss"]:	#物联网四大区HDRA制作内容
			if prov in HDRA_Prov_CH[item]["area"]:
				#BELL
				#Gx指向PCRF，Cx Sh指向IMS-HSS，SLh指向 EPC-HSS
				pcrf_value=HDRA_Prov_CH[item]['pcrf']
				#2019.01.22存量号段不增加Gx
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETAPP=Gx|ROUTESET|{}-rs.\n".format(imsi_value,pcrf_value)
				HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETAPP=Cx|ROUTESET|{}-IMS-rs.\n".format(imsi_value,hss_value)
				HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETAPP=Sh|ROUTESET|{}-IMS-rs.\n".format(imsi_value,hss_value)
				HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETAPP=SLh|ROUTESET|{}-rs.\n".format(imsi_value,hss_value)
				#print (HDRA_Prov_CH[item]['jsj'])
				#华为
				#Cx生成一条IPMU(IMSI) Zh生成一条IMPI（IMSI）S6a生成IMSI
				nextindex=route_hw[HDRA_Prov_CH[item]['name_en']][hss_value]  #指向EPC-HSS
				nextindex_ims=route_hw[HDRA_Prov_CH[item]['name_en']][hss_value+'-IMS']
				nextindex_pcrf=route_hw[HDRA_Prov_CH[item]['name_en']][pcrf_value]
				mog=hss_value
				#mog_ims=hss_value+'-IMS'
				mog_ims=hss_value
				#mog_pcrf=pcrf_value
				mog_pcrf=hss_value
				#2019.01.22  存量号段不生成Gx
				#Cx接口指向IMS-HSS 默认参考值为0
				HW+="ADD RTIMPU: REFERINDEX=0,IMPU=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#Gx接口指向PCRF 默认参考值为0（临时）
				#HW+="ADD RTMSISDN: REFERINDEX=0,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_pcrf,mog_pcrf,HDRA_Prov_CH[item]['name_short'])
				#Sh接口指向IMS-HSS 默认参考值为1！！！
				HW+="ADD RTMSISDN: REFERINDEX=1,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#SLh接口指向EPC-HSS 默认参考值为2！！！
				HW+="ADD RTMSISDN: REFERINDEX=2,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				#Gx接口指向PCRF 默认参考值为3
				#HW+="ADD RTMSISDN: REFERINDEX=3,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_pcrf,mog_pcrf,HDRA_Prov_CH[item]['name_short'])
				############这个判断是否4个大区，4大区采用新的MSISDN接口/上面的if判断的为是否是hss归属大区
				'''if item in HDRA_PROV_WLW:
				#Sh接口指向IMS-HSS 默认参考值为1！！！
					HW+="ADD RTMSISDN: REFERINDEX=1,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#SLh接口指向EPC-HSS 默认参考值为2！！！
					HW+="ADD RTMSISDN: REFERINDEX=2,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				else:
				#Sh接口指向IMS-HSS 默认参考值为0！！！
					HW+="ADD RTMSISDN: REFERINDEX=0,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#SLh接口指向EPC-HSS 默认参考值为0！！！
					HW+="ADD RTMSISDN: REFERINDEX=0,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				'''
			else:
				#BELL
				####S6a指向EPC HSS，现保留原始格式 / Cx Zh指向IMS的HSS
				#Gx指向PCRF，Cx Sh指向IMS-HSS，SLh指向 EPC-HSS
				rt_bl=route_bell[prov]["归属物联网大区"]
				HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETDIR=ROUTESET|{}-rs.\n".format(imsi_value,rt_bl)
				#非号段归属大区只需要配置一条，落地省再区分接口
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETAPP=Gx|ROUTESET|{}-rs.\n".format(imsi_value,rt_bl)
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETAPP=Cx|ROUTESET|{}-rs.\n".format(imsi_value,rt_bl)
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETAPP=Sh|ROUTESET|{}-rs.\n".format(imsi_value,rt_bl)
				#HDRA_Prov_CH[item]['jsj']+="CREATE-DIAM-PROXYDATA:MSISDN=86{},TARGETAPP=SLh|ROUTESET|{}-rs.\n".format(imsi_value,rt_bl)
				
				#华为
				#Cx生成一条IPMU(IMSI) Zh生成一条IMPI（IMSI）S6a生成IMSI
				prov_belong=route_bell[prov]["归属物联网大区"]
				#print (prov_belong)
				nextindex=route_hw[HDRA_Prov_CH[item]['name_en']][prov_belong]

				nextindex=route_hw[HDRA_Prov_CH[item]['name_en']][prov_belong]  #指向归属省份
				nextindex_ims=route_hw[HDRA_Prov_CH[item]['name_en']][prov_belong]
				nextindex_pcrf=route_hw[HDRA_Prov_CH[item]['name_en']][prov_belong]
				#pcrf_value=HDRA_Prov_CH[item]['pcrf']
				mog=hss_value
				#mog_ims=hss_value+'-IMS'
				mog_ims=hss_value
				#mog_pcrf=pcrf_value
				mog_pcrf=hss_value
				#Cx接口指向IMS-HSS 默认参考值为0
				HW+="ADD RTIMPU: REFERINDEX=0,IMPU=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#Gx接口指向PCRF 默认参考值为0
				#HW+="ADD RTMSISDN: REFERINDEX=0,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_pcrf,mog_pcrf,HDRA_Prov_CH[item]['name_short'])
				
				
				############这个判断是否4个大区，4大区采用新的MSISDN接口/上面的if判断的为是否是hss归属大区
				if item in HDRA_PROV_WLW:
				#Sh接口指向IMS-HSS 默认参考值为1！！！
					HW+="ADD RTMSISDN: REFERINDEX=1,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#SLh接口指向EPC-HSS 默认参考值为2！！！
					HW+="ADD RTMSISDN: REFERINDEX=2,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				#Gx接口指向PCRF 默认参考值为0
				#	HW+="ADD RTMSISDN: REFERINDEX=3,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_pcrf,mog_pcrf,HDRA_Prov_CH[item]['name_short'])
				
				#else:
				#Sh接口指向IMS-HSS 默认参考值为0！！！	
				#	HW+="ADD RTMSISDN: REFERINDEX=0,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#SLh接口指向EPC-HSS 默认参考值为0！！！
				#	HW+="ADD RTMSISDN: REFERINDEX=0,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				






				#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>非落地省就用一条MSISDN即可
				######都是非大区省，用默认的0参考号就行
				#Sh接口指向IMS-HSS 默认参考值为1！！！>>>>费大区省参考值为0即可
				#HW+="ADD RTMSISDN: REFERINDEX=0,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex_ims,mog_ims,HDRA_Prov_CH[item]['name_short'])
				#SLh接口指向EPC-HSS 默认参考值为2！！！>>>>费大区省参考值为0即可
				#HW+="ADD RTMSISDN: REFERINDEX=0,MSISDN=\"86{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])





				#HW+="ADD RTIMSI: REFERINDEX=0,IMSI=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				#HW+="ADD RTIMSI: REFERINDEX=0,IMPU=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])
				#HW+="ADD RTIMSI: REFERINDEX=0,IMPI=\"{0}\",NEXTRULE=RTEXIT,NEXTINDEX={1},MOG=\"{2}\";{{{3}DRA01AHW-A}}\n".format(imsi_value,nextindex,mog,HDRA_Prov_CH[item]['name_short'])

	for item in HDRA_Prov_CH.keys():
		c=open(workpath+"\\脚本\\"+"["+HDRA_Prov_CH[item]['name_short']+"DRA01AAL-B]B_"+nowTime+".txt",'w')
		c.write(HDRA_Prov_CH[item]['jsj'])
		c.close()
	c=open(workpath+"\\脚本\\"+"_[HW_HDRA]H_"+nowTime+".txt",'w')
	c.write(HW)
	c.close()


def ROUTE_NEW():
	workbook = xlrd.open_workbook(workpath+"\\CONF\\"+"route.xlsx")
	table=workbook.sheets()[0]
	table1=workbook.sheets()[1]
	route_bell={}
	route_hw={}
	for i in range (1,table.nrows):
		route_bell[table.cell(i,0).value]={}
		route_bell[table.cell(i,0).value]["归属大区"]=table.cell(i,1).value
		route_bell[table.cell(i,0).value]["归属物联网大区"]=table.cell(i,3).value
	for i in range (1,table1.nrows):
		prov=table1.cell(i,0).value
		prov_des=table1.cell(i,1).value
		index=str(int(table1.cell(i,2).value))
		if prov in route_hw.keys():
			route_hw[prov][prov_des]=index
		else:
			route_hw[prov]={}
			route_hw[prov][prov_des]=index
	return (route_bell,route_hw)
		



def main():

	#jsj_name= input("请输入局数据批次：")
	x=run(workpath)

if __name__=="__main__":
	main()
	#input()

