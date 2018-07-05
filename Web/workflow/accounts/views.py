# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import os, csv, io, codecs, xlrd, pythoncom, win32com.client
from django.views.decorators.csrf import ensure_csrf_cookie
from django.views.decorators.csrf import csrf_exempt
from django.http import JsonResponse
from django.http.response import HttpResponse
from django.shortcuts import render, redirect
from datetime import datetime as dt
from django.contrib.auth.models import User
from http.cookiejar import CookieJar
from collections import OrderedDict

login_data = {}
input_dir = os.path.dirname(os.path.abspath(__file__)) + '/static/input/'
output_template_dir = os.path.dirname(os.path.abspath(__file__)) + '/static/output/template/'
output_dir = os.path.dirname(os.path.abspath(__file__)) + '/static/output/'

login_date = {}
def login(req):
    return render(req, 'login.html')

def top(req):
    global login_date
    login_date = { 'datetime':dt.now().strftime("%Y/%m/%d %H:%M") }
    return render(req, 'top.html', login_date)


def helpQA(req):
    global login_date
    return render(req, 'helpQA.html', login_date)

def rgbToInt(rgb):
    colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
    return colorInt
 

@csrf_exempt
def csv_upload(req):
    input_save_path = os.path.join(input_dir, "input.xlsx")
    output_template_path = os.path.join(output_template_dir, "output_template.xlsx")
    output_save_path = os.path.join(output_dir, "workflow.xlsx")

    post_data = []
    if req.method == 'POST':
        post_data = req.FILES['file'].read()
        with open(input_save_path, mode='wb') as f:
            f.write(post_data)

        #inputファイル読み込み(pandasデータフレーム)
        workbook_df = pd.ExcelFile(input_save_path, encoding='utf8')
        sheet_name = workbook_df.sheet_names
        sheet_df = workbook_df.parse(sheet_name[0], header=0)
#        print("Sheet name:", sheet_name[0])
#        print(sheet_df.keys) 

        #outputファイル編集（オートシェイプ描画）
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(output_template_path)
        sheet = workbook.Worksheets(1)
        sheet.Activate()

        #部門設定
        grouped = sheet_df.groupby('部門')
        grouped = sorted(grouped, key=lambda x: len(x[1]), reverse=True)
        grouped_df = pd.DataFrame(index=[], columns=['depart', 'count'])
        for key, group in grouped:
#            print(key, len(group))
            series = pd.Series([key, len(group)], index=grouped_df.columns)
            grouped_df = grouped_df.append(series, ignore_index = True)                 
#        print(df)
        sheet.Cells(3,1).Value = grouped_df.iat[0, 0]
        sheet.Cells(23,1).Value = grouped_df.iat[1, 0]
        sheet.Cells(43,1).Value = grouped_df.iat[2, 0]

#        num_grouped1 = grouped_df.iat[0, 1]
#        num_grouped2 = grouped_df.iat[1, 1]
#        num_grouped3 = grouped_df.iat[2, 1]
#        grouped1_df = (sheet_df.loc[sheet_df['部門'] == grouped_df.iat[0, 0]])
#        print(grouped1_df)
 
        #plot(row[部門、対象、処理])    
        for index, row in sheet_df.iterrows():
            if row[0] == grouped_df.iat[0, 0]:
                if row[2] == '連絡':
                    shape1 = sheet.Shapes.AddShape(18, sheet.cells(1,3).Left, sheet.cells(19,1).Top, 40, 40)
                    shape1.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                elif row[2] == '受取':
                        if row[1] == '見積書':
                            shape6 = sheet.Shapes.AddShape(81,sheet.cells(1,13).Left, sheet.cells(19,1).Top, 40, 40)
                            shape6.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                            text6 = sheet.Shapes.AddShape(1, sheet.cells(1,13).Left, sheet.cells(17,1).Top, 90, 70)
                            text6.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                            text6.TextFrame2.TextRange.Characters.Text = row[1]
                            text6.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                            text6.TextFrame2.TextRange.Characters.Font.Size = 12
                        elif row[1] == '発注書コピー':
                            shape21 = sheet.Shapes.AddShape(81,sheet.cells(1,25).Left, sheet.cells(7,1).Top, 40, 40)
                            shape21.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                            text21 = sheet.Shapes.AddShape(1, sheet.cells(1,25).Left, sheet.cells(5,1).Top, 90, 70)
                            text21.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                            text21.TextFrame2.TextRange.Characters.Text = row[1]
                            text21.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                            text21.TextFrame2.TextRange.Characters.Font.Size = 12
                        elif row[1] == '納品書':
                            shape26 = sheet.Shapes.AddShape(81,sheet.cells(1,35).Left, sheet.cells(19,1).Top, 40, 40)
                            shape26.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                            text26 = sheet.Shapes.AddShape(1, sheet.cells(1,35).Left, sheet.cells(17,1).Top, 90, 70)
                            text26.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                            text26.TextFrame2.TextRange.Characters.Text = row[1]
                            text26.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                            text26.TextFrame2.TextRange.Characters.Font.Size = 12
                        else:
                            continue
                elif row[2] == '確認':
                        if row[1] == '見積書':
                            shape7 = sheet.Shapes.AddShape(4,sheet.cells(1,15).Left, sheet.cells(19,1).Top, 40, 40)
                            shape7.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        elif row[1] == '納品書':
                            shape27 = sheet.Shapes.AddShape(4,sheet.cells(1,37).Left, sheet.cells(19,1).Top, 40, 40)
                            shape27.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        elif row[1] == '現品':
                            shape32 = sheet.Shapes.AddShape(4,sheet.cells(1,33).Left, sheet.cells(16,1).Top, 40, 40)
                            shape32.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        elif row[1] == '発注書コピー（一時保管）':
                            shape34 = sheet.Shapes.AddShape(4,sheet.cells(1,39).Left, sheet.cells(13,1).Top, 40, 40)
                            shape34.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        else:
                            continue
                elif row[2] == '作成':
                    if row[1] == '見積書から仕入申請書':
                        shape9 = sheet.Shapes.AddShape(73,sheet.cells(1,15).Left, sheet.cells(16,1).Top, 40, 40)
                        shape9.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape9.TextFrame2.TextRange.Characters.Text = "F"
                        shape9.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape9.TextFrame2.TextRange.Characters.Font.Size = 22
                        text9 = sheet.Shapes.AddShape(1, sheet.cells(1,15).Left, sheet.cells(14,1).Top, 90, 70)
                        text9.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text9.TextFrame2.TextRange.Characters.Text = row[1]
                        text9.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text9.TextFrame2.TextRange.Characters.Font.Size = 12
                    elif row[1] == '見積書と仕入申請書':
                        shape12 = sheet.Shapes.AddShape(73,sheet.cells(1,19).Left, sheet.cells(13,1).Top, 40, 40)
                        shape12.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape12.TextFrame2.TextRange.Characters.Text = "F"
                        shape12.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape12.TextFrame2.TextRange.Characters.Font.Size = 22
                        text12 = sheet.Shapes.AddShape(1, sheet.cells(1,19).Left, sheet.cells(11,1).Top, 90, 70)
                        text12.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text12.TextFrame2.TextRange.Characters.Text = row[1]
                        text12.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text12.TextFrame2.TextRange.Characters.Font.Size = 12
                    elif row[1] == '見積書と仕入申請書から発注書':
                        shape15 = sheet.Shapes.AddShape(73,sheet.cells(1,21).Left, sheet.cells(10,1).Top, 40, 40)
                        shape15.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape15.TextFrame2.TextRange.Characters.Text = "F"
                        shape15.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape15.TextFrame2.TextRange.Characters.Font.Size = 22
                        text15 = sheet.Shapes.AddShape(1, sheet.cells(1,21).Left, sheet.cells(8,1).Top, 90, 70)
                        text15.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text15.TextFrame2.TextRange.Characters.Text = row[1]
                        text15.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text15.TextFrame2.TextRange.Characters.Font.Size = 12
                    elif row[1] == '見積書と納品書':
                        shape38 = sheet.Shapes.AddShape(73,sheet.cells(1,43).Left, sheet.cells(7,1).Top, 40, 40)
                        shape38.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape38.TextFrame2.TextRange.Characters.Text = "F"
                        shape38.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape38.TextFrame2.TextRange.Characters.Font.Size = 22
                        text38 = sheet.Shapes.AddShape(1, sheet.cells(1,43).Left, sheet.cells(5,1).Top, 90, 70)
                        text38.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text38.TextFrame2.TextRange.Characters.Text = row[1]
                        text38.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text38.TextFrame2.TextRange.Characters.Font.Size = 12
                elif row[2] == '記入':
                    if row[1] == '担当者が仕入申請書':
                        shape10 = sheet.Shapes.AddShape(73, sheet.cells(1,17).Left, sheet.cells(16,1).Top, 40, 40)
                        shape10.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape10.TextFrame2.TextRange.Characters.Text = "F"
                        shape10.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape10.TextFrame2.TextRange.Characters.Font.Size = 22
                    elif row[1] == '担当者が発注書':
                        shape16 = sheet.Shapes.AddShape(73, sheet.cells(1,23).Left, sheet.cells(10,1).Top, 40, 40)
                        shape16.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape16.TextFrame2.TextRange.Characters.Text = "F"
                        shape16.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape16.TextFrame2.TextRange.Characters.Font.Size = 22
                    elif row[1] == '納品書':
                        shape28 =sheet.Shapes.AddShape(73,sheet.cells(1,39).Left, sheet.cells(19,1).Top, 40, 40)
                        shape28.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape28.TextFrame2.TextRange.Characters.Text = "F"
                        shape28.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape28.TextFrame2.TextRange.Characters.Font.Size = 22
                    else:
                        continue                    
                elif row[2] == '集める':
                    if row[1] == '見積書':                    
                        shape8 =sheet.Shapes.AddShape(77,sheet.cells(1,19).Left, sheet.cells(19,1).Top, 40, 40)
                        shape8.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    elif row[1] == '仕入申請書':
                        shape11 = sheet.Shapes.AddShape(77,sheet.cells(1,19).Left, sheet.cells(16,1).Top, 40, 40)
                        shape11.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    elif row[1] == '納品書':
                        shape29 =sheet.Shapes.AddShape(77,sheet.cells(1,43).Left, sheet.cells(19,1).Top, 40, 40)
                        shape29.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    elif row[1] == '見積書（一時保管）':
                        shape37 =sheet.Shapes.AddShape(77,sheet.cells(1,43).Left, sheet.cells(10,1).Top, 40, 40)
                        shape37.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                elif row[2] == '承認':
                    shape13 = sheet.Shapes.AddShape(73,sheet.cells(1,21).Left, sheet.cells(13,1).Top, 40, 40)
                    shape13.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    shape13.TextFrame2.TextRange.Characters.Text = "S"
                    shape13.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                    shape13.TextFrame2.TextRange.Characters.Font.Size = 22
                elif row[2] == '一時保管':
                    if row[1] == '見積書と仕入申請書':
                        shape14 = sheet.Shapes.AddShape(82, sheet.cells(1,23).Left, sheet.cells(13,1).Top, 40, 40)
                        shape14.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    elif row[1] == '発注書コピー':
                        shape22 = sheet.Shapes.AddShape(82, sheet.cells(1,27).Left, sheet.cells(7,1).Top, 40, 40)
                        shape22.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    elif row[1] == '発注書コピー（一時保管）':
                        shape35 = sheet.Shapes.AddShape(82,sheet.cells(1,41).Left, sheet.cells(13,1).Top, 40, 40)
                        shape35.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                elif row[2] == '郵送':
                    shape17 = sheet.Shapes.AddShape(73,sheet.cells(1,25).Left, sheet.cells(10,1).Top, 40, 40)
                    shape17.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    shape17.TextFrame2.TextRange.Characters.Text = "〒"
                    shape17.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                    shape17.TextFrame2.TextRange.Characters.Font.Size = 16
                elif row[2] == 'コピー':
                    shape20 = sheet.Shapes.AddShape(73,sheet.cells(1,23).Left, sheet.cells(7,1).Top, 40, 40)
                    shape20.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    shape20.TextFrame2.TextRange.Characters.Text = "C"
                    shape20.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                    shape20.TextFrame2.TextRange.Characters.Font.Size = 22
                    text20 = sheet.Shapes.AddShape(1, sheet.cells(1,23).Left, sheet.cells(5,1).Top, 90, 70)
                    text20.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    text20.TextFrame2.TextRange.Characters.Text = row[1]
                    text20.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                    text20.TextFrame2.TextRange.Characters.Font.Size = 12
                elif row[2] == '取り出す':
                    if row[1] == '発注書コピー（一時保管）':
                        shape33 = sheet.Shapes.AddShape(82,sheet.cells(1,37).Left, sheet.cells(13,1).Top, 40, 40)
#                        shape33.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                    if row[1] == '見積書（一時保管）':
                        shape36 = sheet.Shapes.AddShape(82,sheet.cells(1,39).Left, sheet.cells(10,1).Top, 40, 40)
#                        shape36.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                elif row[2] == '運ぶ':
                        shape39 = sheet.Shapes.AddShape(73,sheet.cells(1,45).Left, sheet.cells(7,1).Top, 40, 40)
                        shape39.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape39.TextFrame2.TextRange.Characters.Text = "P"
                        shape39.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape39.TextFrame2.TextRange.Characters.Font.Size = 22

            elif row[0] == grouped_df.iat[1, 0]:
                if row[2] == '連絡':
                    shape2 = sheet.Shapes.AddShape(18,sheet.cells(1,5).Left, sheet.cells(39,1).Top, 40, 40)
                    shape2.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                elif row[2] == '作成':
                    if row[1] == '見積書':                    
                        shape3 = sheet.Shapes.AddShape(73,sheet.cells(1,7).Left, sheet.cells(39,1).Top, 40, 40)
                        shape3.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape3.TextFrame2.TextRange.Characters.Text = "F"
                        shape3.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape3.TextFrame2.TextRange.Characters.Font.Size = 22
                    elif row[1] == '発注書から納品書':
                        shape23 = sheet.Shapes.AddShape(73,sheet.cells(1,29).Left, sheet.cells(33,1).Top, 40, 40)
                        shape23.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape23.TextFrame2.TextRange.Characters.Text = "F"
                        shape23.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape23.TextFrame2.TextRange.Characters.Font.Size = 22
                        text23 = sheet.Shapes.AddShape(1, sheet.cells(1,29).Left, sheet.cells(31,1).Top, 90, 70)
                        text23.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text23.TextFrame2.TextRange.Characters.Text = row[1]
                        text23.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text23.TextFrame2.TextRange.Characters.Font.Size = 12
                elif row[2] == '省略':
                    if row[1] == '見積書':                    
                        shape4 = sheet.Shapes.AddShape(73,sheet.cells(1,9).Left, sheet.cells(39,1).Top, 40, 40)
                        shape4.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape4.TextFrame2.TextRange.Characters.Text = "E"
                        shape4.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape4.TextFrame2.TextRange.Characters.Font.Size = 22
                    elif row[1] == '発注書':
                        shape19 = sheet.Shapes.AddShape(73,sheet.cells(1,29).Left, sheet.cells(36,1).Top, 40, 40)
                        shape19.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape19.TextFrame2.TextRange.Characters.Text = "E"
                        shape19.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape19.TextFrame2.TextRange.Characters.Font.Size = 22                        
                    elif row[1] == '納品書':
                        shape24 = sheet.Shapes.AddShape(73, sheet.cells(1,31).Left, sheet.cells(33,1).Top, 40, 40)
                        shape24.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape24.TextFrame2.TextRange.Characters.Text = "E"
                        shape24.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape24.TextFrame2.TextRange.Characters.Font.Size = 22
                elif row[2] == '郵送':
                    if row[1] == '見積書':
                        shape5 = sheet.Shapes.AddShape(73, sheet.cells(1,11).Left, sheet.cells(39,1).Top, 40, 40)
                        shape5.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape5.TextFrame2.TextRange.Characters.Text = "〒"
                        shape5.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape5.TextFrame2.TextRange.Characters.Font.Size = 16
                    elif row[1] == '納品書':
                        shape25 = sheet.Shapes.AddShape(73, sheet.cells(1,33).Left, sheet.cells(33,1).Top, 40, 40)
                        shape25.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape25.TextFrame2.TextRange.Characters.Text = "〒"
                        shape25.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape25.TextFrame2.TextRange.Characters.Font.Size = 16
                    elif row[1] == '現品':
                        shape31 = sheet.Shapes.AddShape(73,sheet.cells(1,33).Left, sheet.cells(30,1).Top, 40, 40)
                        shape31.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape31.TextFrame2.TextRange.Characters.Text = "〒"
                        shape31.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape31.TextFrame2.TextRange.Characters.Font.Size = 16
                elif row[2] == '受取':
                    shape18 = sheet.Shapes.AddShape(81, sheet.cells(1,27).Left, sheet.cells(36,1).Top, 40, 40)
                    shape18.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    text18 = sheet.Shapes.AddShape(1, sheet.cells(1,27).Left, sheet.cells(34,1).Top, 90, 70)
                    text18.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    text18.TextFrame2.TextRange.Characters.Text = row[1]
                    text18.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                    text18.TextFrame2.TextRange.Characters.Font.Size = 12
                elif row[2] == '作る':
                        shape30 = sheet.Shapes.AddShape(14, sheet.cells(1,31).Left, sheet.cells(30,1).Top, 50, 50)
                        shape30.Fill.ForeColor.RGB = rgbToInt((255,255,255))

            elif row[0] == grouped_df.iat[2, 0]:
                if row[2] == '受取':
                    shape40 = sheet.Shapes.AddShape(81,sheet.cells(1,47).Left, sheet.cells(59,1).Top, 40, 40)
                    shape40.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    text40 = sheet.Shapes.AddShape(1, sheet.cells(1,47).Left, sheet.cells(57,1).Top, 90, 70)
                    text40.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                    text40.TextFrame2.TextRange.Characters.Text = row[1]
                    text40.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                    text40.TextFrame2.TextRange.Characters.Font.Size = 12
                if row[2] == '保管':
                    shape41 = sheet.Shapes.AddShape(82,sheet.cells(1,49).Left, sheet.cells(59,1).Top, 40, 40)
                    shape41.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                if row[2] == '立ち上げる':
                    if row[1] == '支払依頼画面':
                        shape42 = sheet.Shapes.AddShape(73,sheet.cells(1,49).Left, sheet.cells(56,1).Top, 40, 40)
                        shape42.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape42.TextFrame2.TextRange.Characters.Text = "㍶"
                        shape42.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape42.TextFrame2.TextRange.Characters.Font.Size = 16                        
                        text42 = sheet.Shapes.AddShape(1, sheet.cells(1,49).Left, sheet.cells(54,1).Top, 90, 70)
                        text42.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text42.TextFrame2.TextRange.Characters.Text = row[1]
                        text42.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text42.TextFrame2.TextRange.Characters.Font.Size = 12
                    elif row[1] == '買掛金画面':
                        shape44 = sheet.Shapes.AddShape(73,sheet.cells(1,51).Left, sheet.cells(53,1).Top, 40, 40)
                        shape44.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape44.TextFrame2.TextRange.Characters.Text = "㍶"
                        shape44.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape44.TextFrame2.TextRange.Characters.Font.Size = 16
                        text44 = sheet.Shapes.AddShape(1, sheet.cells(1,51).Left, sheet.cells(51,1).Top, 90, 70)
                        text44.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text44.TextFrame2.TextRange.Characters.Text = row[1]
                        text44.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text44.TextFrame2.TextRange.Characters.Font.Size = 12
                if row[2] == '入力':
                    if row[1] == '見積書と納品書から支払依頼データ':
                        shape43 = sheet.Shapes.AddShape(73,sheet.cells(1,51).Left, sheet.cells(56,1).Top, 40, 40)
                        shape43.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape43.TextFrame2.TextRange.Characters.Text = "K"
                        shape43.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape43.TextFrame2.TextRange.Characters.Font.Size = 22
                        text43 = sheet.Shapes.AddShape(1, sheet.cells(1,51).Left, sheet.cells(54,1).Top, 90, 70)
                        text43.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text43.TextFrame2.TextRange.Characters.Text = row[1]
                        text43.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text43.TextFrame2.TextRange.Characters.Font.Size = 12
                    elif row[1] == '見積書と納品書から買掛金データ':
                        shape45 = sheet.Shapes.AddShape(73,sheet.cells(1,53).Left, sheet.cells(53,1).Top, 40, 40)
                        shape45.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        shape45.TextFrame2.TextRange.Characters.Text = "K"
                        shape45.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        shape45.TextFrame2.TextRange.Characters.Font.Size = 22
                        text45 = sheet.Shapes.AddShape(1, sheet.cells(1,53).Left, sheet.cells(51,1).Top, 90, 70)
                        text45.Fill.ForeColor.RGB = rgbToInt((255,255,255))
                        text45.TextFrame2.TextRange.Characters.Text = row[1]
                        text45.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
                        text45.TextFrame2.TextRange.Characters.Font.Size = 12

#        #コネクターの書き方　
        connector1 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector1.ConnectorFormat.BeginConnect(shape1, 5)
        connector1.ConnectorFormat.EndConnect(shape2, 1)
        connector1.Line.EndArrowheadStyle = 4
        connector1.Line.Weight = 2
        connector1.Line.EndArrowheadWidth = 3
        connector1.Line.EndArrowheadLength = 3

        connector2 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector2.ConnectorFormat.BeginConnect(shape2, 7)
        connector2.ConnectorFormat.EndConnect(shape3, 3)
        connector2.Line.Weight = 2
        
        connector3 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector3.ConnectorFormat.BeginConnect(shape3, 7)
        connector3.ConnectorFormat.EndConnect(shape4, 3)
        connector3.Line.Weight = 2
        
        connector4 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector4.ConnectorFormat.BeginConnect(shape4, 7)
        connector4.ConnectorFormat.EndConnect(shape5, 3)
        connector4.Line.Weight = 2
        
        connector5 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector5.ConnectorFormat.BeginConnect(shape5, 1)
        connector5.ConnectorFormat.EndConnect(shape6, 3)
        connector5.Line.EndArrowheadStyle = 4
        connector5.Line.Weight = 2
        connector5.Line.EndArrowheadWidth = 3
        connector5.Line.EndArrowheadLength = 3
        
        connector6 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector6.ConnectorFormat.BeginConnect(shape6, 4)
        connector6.ConnectorFormat.EndConnect(shape7, 2)
        connector6.Line.Weight = 2
        
        
        connector7 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector7.ConnectorFormat.BeginConnect(shape7, 4)
        connector7.ConnectorFormat.EndConnect(shape8, 3)
        connector7.Line.Weight = 2
        
        connector8 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector8.ConnectorFormat.BeginConnect(shape10, 7)
        connector8.ConnectorFormat.EndConnect(shape11, 3)
        connector8.Line.Weight = 2
        
        connector9 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector9.ConnectorFormat.BeginConnect(shape9, 7)
        connector9.ConnectorFormat.EndConnect(shape10, 3)
        connector9.Line.Weight = 2
        
        connector10 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector10.ConnectorFormat.BeginConnect(shape12, 7)
        connector10.ConnectorFormat.EndConnect(shape13, 3)
        connector10.Line.Weight = 2
        
        connector11 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector11.ConnectorFormat.BeginConnect(shape13, 7)
        connector11.ConnectorFormat.EndConnect(shape14, 2)
        connector11.Line.Weight = 2
        
        connector12 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector12.ConnectorFormat.BeginConnect(shape15, 7)
        connector12.ConnectorFormat.EndConnect(shape16, 3)
        connector12.Line.Weight = 2
        
        connector13 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector13.ConnectorFormat.BeginConnect(shape16, 7)
        connector13.ConnectorFormat.EndConnect(shape17, 3)
        connector13.Line.Weight = 2
        
        connector14 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector14.ConnectorFormat.BeginConnect(shape20, 7)
        connector14.ConnectorFormat.EndConnect(shape21, 2)
        connector14.Line.Weight = 2
        
        connector15 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector15.ConnectorFormat.BeginConnect(shape21, 4)
        connector15.ConnectorFormat.EndConnect(shape22, 2)
        connector15.Line.Weight = 2
        
        connector16 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector16.ConnectorFormat.BeginConnect(shape17, 5)
        connector16.ConnectorFormat.EndConnect(shape18, 2)
        connector16.Line.EndArrowheadStyle = 4
        connector16.Line.Weight = 2
        connector16.Line.EndArrowheadWidth = 3
        connector16.Line.EndArrowheadLength = 3

        connector17 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector17.ConnectorFormat.BeginConnect(shape18, 4)
        connector17.ConnectorFormat.EndConnect(shape19, 3)
        connector17.Line.Weight = 2
        
        connector18 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector18.ConnectorFormat.BeginConnect(shape23, 7)
        connector18.ConnectorFormat.EndConnect(shape24, 3)
        connector18.Line.Weight = 2
        
        connector19 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector19.ConnectorFormat.BeginConnect(shape24, 7)
        connector19.ConnectorFormat.EndConnect(shape25, 3)
        connector19.Line.Weight = 2
        
        connector20 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector20.ConnectorFormat.BeginConnect(shape26, 4)
        connector20.ConnectorFormat.EndConnect(shape27, 2)
        connector20.Line.Weight = 2
        
        connector21 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector21.ConnectorFormat.BeginConnect(shape25, 7)
        connector21.ConnectorFormat.EndConnect(shape26, 3)
        connector21.Line.EndArrowheadStyle = 4
        connector21.Line.Weight = 2
        connector21.Line.EndArrowheadWidth = 3
        connector21.Line.EndArrowheadLength = 3
        
        connector22 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector22.ConnectorFormat.BeginConnect(shape31, 1)
        connector22.ConnectorFormat.EndConnect(shape32, 3)
        connector22.Line.EndArrowheadStyle = 4
        connector22.Line.Weight = 2
        connector22.Line.EndArrowheadWidth = 3
        connector22.Line.EndArrowheadLength = 3
        
        connector23 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector23.ConnectorFormat.BeginConnect(shape27, 4)
        connector23.ConnectorFormat.EndConnect(shape28, 3)
        connector23.Line.Weight = 2
        
        connector24 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector24.ConnectorFormat.BeginConnect(shape28, 7)
        connector24.ConnectorFormat.EndConnect(shape29, 3)
        connector24.Line.Weight = 2
        
        connector25 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector25.ConnectorFormat.BeginConnect(shape33, 4)
        connector25.ConnectorFormat.EndConnect(shape34, 2)
        connector25.Line.Weight = 2
        
        connector26 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector26.ConnectorFormat.BeginConnect(shape30, 6)
        connector26.ConnectorFormat.EndConnect(shape31, 3)
        connector26.Line.Weight = 2
        
        connector27 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector27.ConnectorFormat.BeginConnect(shape34, 4)
        connector27.ConnectorFormat.EndConnect(shape35, 2)
        connector27.Line.Weight = 2
        
        connector28 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector28.ConnectorFormat.BeginConnect(shape36, 4)
        connector28.ConnectorFormat.EndConnect(shape37, 3)
        connector28.Line.Weight = 2
        
        connector29 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector29.ConnectorFormat.BeginConnect(shape38, 7)
        connector29.ConnectorFormat.EndConnect(shape39, 3)
        connector29.Line.Weight = 2
        
        connector30 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector30.ConnectorFormat.BeginConnect(shape39, 5)
        connector30.ConnectorFormat.EndConnect(shape40, 2)
        connector30.Line.EndArrowheadStyle = 4
        connector30.Line.Weight = 2
        connector30.Line.EndArrowheadWidth = 3
        connector30.Line.EndArrowheadLength = 3
        
        connector31 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector31.ConnectorFormat.BeginConnect(shape40, 4)
        connector31.ConnectorFormat.EndConnect(shape41, 2)
        connector31.Line.Weight = 2
        
        connector32 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector32.ConnectorFormat.BeginConnect(shape42, 7)
        connector32.ConnectorFormat.EndConnect(shape43, 3)
        connector32.Line.Weight = 2
        
        connector33 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector33.ConnectorFormat.BeginConnect(shape44, 7)
        connector33.ConnectorFormat.EndConnect(shape45, 3)
        connector33.Line.Weight = 2


        
        workbook.SaveAs(output_save_path) #saveas名前をつけて保存
        workbook.Close()
        excel.Quit()
        sheet = None
        workbook = None
        excel = None

    return render(req, 'top.html', login_date)

@csrf_exempt
def csv_delete(req):
    input_path = os.path.join(input_dir, "input.xlsx")

    if input_path:
      os.remove(input_path)
    else:
      print(input_path)
      
    return render(req, 'top.html', login_data)

