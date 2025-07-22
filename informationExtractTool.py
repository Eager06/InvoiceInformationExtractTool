import pdfplumber
import pandas as pd
import re
import openpyxl as oxl
#C:\Users\15233\Desktop\invoices\
#输入文件夹路径：
path = input("Please enter the path of the folder:")

#遍历文件夹获取文件名并加入列表
import os 
file_name_list = []

for dirpath, dirnames, filenames in os.walk(path):
    for filename in filenames:
        if filename.lower().endswith(".pdf"):
            file_name_list.append(filename)
    
         



####创建与修改xlsx表格：
wb = oxl.Workbook()
sheet = wb.active
sheet.title = "发票提取信息"
#设置表头：
headers = ["发票种类", "发票号码", "开票日期", "单位名称（销售方）", "开票申领人", "备注", 
        "文件名", "商品编码", "货物或应税劳务、服务名称", "规格型号", "单位", 
        "数量", "单价", "金额", "税率", "税额", "价税合计", "单号", 
        "回款日期", "回款金额", "买家账户"]
    #填充表头：
for col_num, header in enumerate(headers, 1):
    sheet.cell(row = 1, column = col_num, value = header)


text = ""
Text = ""
#建立字典储存信息：
def dic():
    information_dict = {
        "发票种类" : "无",
        "发票号码" : "无",
        "开票日期" : "无",
        "单位名称（销售方）" : "无",
        "备注" : "无",
        "文件名" : "无",
        "商品编码" : "无",
        "货物或应税劳务、服务名称" : "无",
        "单位" : "无",
        "数量" : "无",
        "单价" : "无",
        "金额" : "无",
        "税率" : "无",
        "税额" : "无",
        "价税合计" : "无",
        "单号" : "无",
        "回款日期" : "无",
        "回款金额" : "无",
        "买家账户" : "无", 
    }

num = 0
#建立列表记录项目数量：
num_of_proj_list = []
kind_list = []
invoice_num_list = []
invoice_date_list = []
name_of_seller_list = []
invoice_giver_list = []
item_code_list = []
TY_name_list = []


for file in file_name_list:

    with pdfplumber.open(path+file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()     #提取文字
            Text += (page.extract_text()+"\n")

            ###开始数据提取
            ##提取文件名：
            pattern_proj = r"\*"
            matches_proj = re.findall(pattern_proj, text)
            num_of_proj = int(len(matches_proj)/2)
            num_of_proj_list.append(num_of_proj)    #用于判断一张发票占几行
            num_of_proj = 0
            ##提取发票种类：
            pattern_kind = r"电子发票\s*（([^）]+)）"
            match_kind = re.search(pattern_kind, text)
            kind_list.append(match_kind.group(1))
            ##提取发票号码：
            pattern_invoice_num = r"发票号码：(\w+)"
            match_invoice_num = re.search(pattern_invoice_num, text)
            invoice_num_list.append(match_invoice_num.group(1))
            ##提取开票日期
            pattern_invoice_date = r"开票日期：(\w+)"
            match_invoice_date = re.search(pattern_invoice_date, text)
            invoice_date_list.append(match_invoice_date.group(1))
            ##提取销售方单位名称
            pattern_name_of_seller = r"销 名称：(\w+)"
            match_name_of_seller = re.search(pattern_name_of_seller, text)
            name_of_seller_list.append(match_name_of_seller.group(1))
            ##提取开票申领人
            pattern_invoice_giver = r"销 名称：(\w+)"
            match_invoice_giver = re.search(pattern_invoice_giver, text)
            invoice_giver_list.append(match_invoice_giver.group(1))
            ##提取商品编码
            pattern_item_code = r"\*(.*?)\*"
            match_item_code = re.search(pattern_item_code,text)
            item_code_list.append(match_item_code.group(1))
            ##提取货物或应税劳务、服务名称
            pattern_TY_name = r"\*.*?\*\s*([^\s]*)"
            match_TY_name = re.search(pattern_TY_name,text)
            TY_name_list.append(match_TY_name.group(1))
            ##识别项数
            pattern_GGXH = r"规格型号"
            if re.search(pattern_GGXH, text):
                num = 8
                print("有规格型号")
            else:
                num = 6
                print("没有规格型号")
            ##识别数据数
            lines = text.splitlines()
            for idx, line in enumerate(text):
                if '*' in line:
                    space_count = line.count(' ')
                    print(f"第{idx + 1}行有星号，共有空格数：{space_count}")





            

#print(TY_name_list)




#C:\Users\15233\Desktop\invoices\

with open(r"C:\Users\15233\Desktop\data.txt", "w", encoding = "utf-8") as file:
    file.write(Text+"\n")


#将文件名添加到表格中
# 将文件名按 num_of_proj_list 中的数量插入 Excel 表格中
row = 2  # 从第二行开始插入数据
for i, num_of_proj in enumerate(num_of_proj_list):
    file_name = file_name_list[i]
    for _ in range(num_of_proj):
        sheet[f"G{row}"] = file_name  # 将文件名插入 G 列
        sheet[f"A{row}"] = kind_list[i-1]
        sheet[f"B{row}"] = invoice_num_list[i-1]
        sheet[f"C{row}"] = invoice_date_list[i-1]
        sheet[f"D{row}"] = name_of_seller_list[i-1]
        sheet[f"E{row}"] = invoice_giver_list[i-1]
        sheet[f"H{row}"] = item_code_list[i-1]
        sheet[f"I{row}"] = TY_name_list[i-1]
        row += 1  # 增加行号，插入下一行






    
#保存文件：   
wb.save(r"C:\Users\15233\Desktop\invoice_data.xlsx")







            















    






