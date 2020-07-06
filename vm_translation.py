'''
MV翻译姬
'''
# -*- coding: utf-8 -*-
import os
import json
import xlwt
import xlrd
import time



def sheetcreation(filename):
	global workbook
	worksheet = workbook.add_sheet(filename)
	worksheet.write(0, 0, "原文")
	worksheet.write(0, 1, "译文")
	return worksheet


if __name__=="__main__":
	selecttype=input("导出文本到excel，请输1；导入文本到游戏，请输2")
	#导入项
	if selecttype==str(1):
		filenum=1
		workbook = xlwt.Workbook()
		worksheet = workbook.add_sheet("sheet1")
		workbook.save("翻译姬.xls")
		rootdir = 'data'
		list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
		for i in range(0, len(list)):
			path = os.path.join(rootdir, list[i])
			filename=list[i]
			if filename.endswith(".json")==True:
				try:
					f = open("data\\"+filename, "r",encoding="utf-8")
					cworksheet=sheetcreation(filename)
					evdata=f.read()
					evdata=json.loads(evdata)
					textnum=1
					for n in range(len(evdata["events"])):
						try:
							for m in range(len(evdata["events"][n]["pages"])):
								for p in range(len(evdata["events"][n]["pages"][m]["list"])):
									if evdata["events"][n]["pages"][m]["list"][p]["code"] == 401:
										cworksheet.write(textnum, 0,evdata["events"][n]["pages"][m]["list"][p]["parameters"][0])
										cworksheet.write(textnum, 1,evdata["events"][n]["pages"][m]["list"][p]["parameters"][0])
										textnum = textnum + 1
									if evdata["events"][n]["pages"][m]["list"][p]["code"] == 102:
										for q in range(len(evdata["events"][n]["pages"][m]["list"][p]["parameters"][0])):
											cworksheet.write(textnum, 0, evdata["events"][n]["pages"][m]["list"][p]["parameters"][0][q])
											cworksheet.write(textnum, 1, evdata["events"][n]["pages"][m]["list"][p]["parameters"][0][q])
											textnum=textnum+1
						except:
							pass
					print("成功读取文件{}个～".format(filenum))
					filenum=filenum+1
					workbook.save("翻译姬.xls")
					f.close()
				except:
					pass
		print("导出文本到excel完成～")
	#导出项
	if selecttype==str(2):
		filenum = 1
		workbook = xlrd.open_workbook("翻译姬.xls")
		rootdir = 'data'
		list = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
		for i in range(0, len(list)):
			time.sleep(1)
			path = os.path.join(rootdir, list[i])
			filename = list[i]
			if filename.endswith(".json") == True:
				try:
					print("正在导出:" + filename)
					f1 = open("data\\" + filename, "r", encoding="utf-8")
					cworksheet = table = workbook.sheet_by_name(filename)
					evdata = f1.read()
					evdata = json.loads(evdata)
					textnum=0
					for n in range(len(evdata["events"])):
						try:
							for m in range(len(evdata["events"][n]["pages"])):
								for p in range(len(evdata["events"][n]["pages"][m]["list"])):
									if evdata["events"][n]["pages"][m]["list"][p]["code"] == 401:
										for r in range(cworksheet.nrows):
											if str(evdata["events"][n]["pages"][m]["list"][p]["parameters"][0])== str(cworksheet.cell(r, 0).value):
												evdata["events"][n]["pages"][m]["list"][p]["parameters"][0] =(str(cworksheet.cell(r, 1).value))
												textnum=1+textnum
									if evdata["events"][n]["pages"][m]["list"][p]["code"] == 102:
										for q in range(len(evdata["events"][n]["pages"][m]["list"][p]["parameters"][0])):
											for r in range(cworksheet.nrows):
												if str(evdata["events"][n]["pages"][m]["list"][p]["parameters"][0][q]) ==str(cworksheet.cell(r, 0).value):
													evdata["events"][n]["pages"][m]["list"][p]["parameters"][0][q] = (str(cworksheet.cell(r, 1).value))
													textnum = 1 + textnum
						except:
							pass

					print("导入文件成功{}个～".format(filenum))
					filenum = filenum + 1
					evdata=json.dumps(evdata)
					f2 = open("data\\" + "temp", "w", encoding="utf-8")
					f2.write(evdata)
					f2.close()
					f1.close()
					os.remove("data\\" + filename)
					os.rename("data\\" + "temp", "data\\" + filename)
				except:
					pass
		print("导入文本到游戏完成～")
	#其它项
	if (selecttype!=str(1))&(selecttype!=str(2)):
		print("输入值错误，请输入1或2")




