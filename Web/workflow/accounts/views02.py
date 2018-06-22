# -*- coding: utf-8 -*-
import os, sys, cv2, traceback, json, requests, time, csv, io, codecs, urllib3
from django.views.decorators.csrf import ensure_csrf_cookie
from django.views.decorators.csrf import csrf_exempt
from django.http import JsonResponse
from django.http.response import HttpResponse
from django.shortcuts import render, redirect
from datetime import datetime as dt
from django.contrib.auth.models import User
from urllib.request import urlopen, Request, build_opener, HTTPCookieProcessor, urlretrieve
from urllib.parse import quote, urlencode, parse_qs
from http.cookiejar import CookieJar
from mimetypes import guess_extension
from time import sleep
from bs4 import BeautifulSoup
import pythoncom, win32com.client, threading

login_data = {}
input_dir = os.path.dirname(os.path.abspath(__file__)) + '/input/'
template_dir = os.path.dirname(os.path.abspath(__file__)) + '/output/template/'
output_dir = os.path.dirname(os.path.abspath(__file__)) + '/output/'

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
    input_path = os.path.join(input_dir, "input.csv")
    template_path = os.path.join(template_dir, "output_template.xlsx")
    output_save_path = os.path.join(output_dir, "output.xlsx")
#    if req.method == 'POST':
#       csv_file = req.FILES['file']
#    
#    with open(csv_file, 'r') as f:
#         reader = csv.reader(f) 
#         for row in reader:
#    
#             print(row) 
#   
    post_data = []
    csv_file = []
    if req.method == 'POST':
        post_data = io.TextIOWrapper(req.FILES['file'])
        csv_file = csv.reader(post_data)
        header = next(csv_file)
        for row in csv_file:
            if row:
                print(row[0])
            else:
                break

        with io.open(input_path, 'w', newline='', encoding='utf-16') as f:
            writer = csv.writer(f, delimiter='t', lineterminator='\n')
            writer.writerows(csv_file)
            print("input_file save complete")
            
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(template_path)
        sheet = workbook.Worksheets(1)
#       sheet = workbook.Sheets('Sheet1').Select(); 
        sheet.Activate()
        sheet.Cells(1,1).Value="購買部"
        sheet.Cells(16,1).Value="仕入先"
        sheet.Cells(31,1).Value="経理部"
        
        

        
        shape1 = sheet.Shapes.AddShape(18,sheet.cells(1,2).Left, sheet.cells(8,1).Top, 30, 30)
        shape1.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape1.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        
        shape2 = sheet.Shapes.AddShape(18,sheet.cells(1,3).Left, sheet.cells(23,1).Top, 30, 30)
        shape2.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape2.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        #Ⓕの書き方
        shape3 = sheet.Shapes.AddShape(73,sheet.cells(1,4).Left, sheet.cells(23,1).Top, 30, 30)
        shape3.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape3.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape3.TextFrame2.TextRange.Characters.Text = "F"
        shape3.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape3.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape4 = sheet.Shapes.AddShape(73,sheet.cells(1,5).Left, sheet.cells(23,1).Top, 30, 30)
        shape4.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape4.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape4.TextFrame2.TextRange.Characters.Text = "E"
        shape4.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape4.TextFrame2.TextRange.Characters.Font.Size = 8
        
        #〶の書き方
        shape5 = sheet.Shapes.AddShape(73,sheet.cells(1,6).Left, sheet.cells(23,1).Top, 30, 30)
        shape5.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape5.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape5.TextFrame2.TextRange.Characters.Text = "〒"
        shape5.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape5.TextFrame2.TextRange.Characters.Font.Size = 8
        
        #△の書き方
        shape6 = sheet.Shapes.AddShape(81,sheet.cells(1,7).Left, sheet.cells(9,1).Top, 30, 30)
        shape6.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape6.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape7 = sheet.Shapes.AddShape(4,sheet.cells(1,8).Left, sheet.cells(9,1).Top, 30, 30)
        shape7.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape7.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape8 = sheet.Shapes.AddShape(73,sheet.cells(1,9).Left, sheet.cells(7,1).Top, 30, 30)
        shape8.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape8.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape8.TextFrame2.TextRange.Characters.Text = "F"
        shape8.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape8.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape9 = sheet.Shapes.AddShape(73,sheet.cells(1,10).Left, sheet.cells(7,1).Top, 30, 30)
        shape9.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape9.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape9.TextFrame2.TextRange.Characters.Text = "F"
        shape9.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape9.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape10 =sheet.Shapes.AddShape(77,sheet.cells(1,11).Left, sheet.cells(9,1).Top, 30, 30)
        shape10.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape10.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape11 = sheet.Shapes.AddShape(77,sheet.cells(1,11).Left, sheet.cells(7,1).Top, 30, 30)
        shape11.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape11.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape12 = sheet.Shapes.AddShape(73,sheet.cells(1,11).Left, sheet.cells(6,1).Top, 30, 30)
        shape12.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape12.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape12.TextFrame2.TextRange.Characters.Text = "F"
        shape12.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape12.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape13 = sheet.Shapes.AddShape(73,sheet.cells(1,12).Left, sheet.cells(6,1).Top, 30, 30)
        shape13.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape13.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape13.TextFrame2.TextRange.Characters.Text = "S"
        shape13.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape13.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape14 = sheet.Shapes.AddShape(73,sheet.cells(1,13).Left, sheet.cells(5,1).Top, 30, 30)
        shape14.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape14.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape14.TextFrame2.TextRange.Characters.Text = "F"
        shape14.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape14.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape15 = sheet.Shapes.AddShape(73,sheet.cells(1,14).Left, sheet.cells(5,1).Top, 30, 30)
        shape15.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape15.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape15.TextFrame2.TextRange.Characters.Text = "F"
        shape15.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape15.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape16 = sheet.Shapes.AddShape(73,sheet.cells(1,15).Left, sheet.cells(4,1).Top, 30, 30)
        shape16.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape16.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape16.TextFrame2.TextRange.Characters.Text = "C"
        shape16.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape16.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape17 = sheet.Shapes.AddShape(82,sheet.cells(1,15).Left, sheet.cells(6,1).Top, 30, 30)
        shape17.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape17.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape18 = sheet.Shapes.AddShape(73,sheet.cells(1,16).Left, sheet.cells(5,1).Top, 30, 30)
        shape18.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape18.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape18.TextFrame2.TextRange.Characters.Text = "〒"
        shape18.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape18.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape19 = sheet.Shapes.AddShape(81,sheet.cells(1,17).Left, sheet.cells(23,1).Top, 30, 30)
        shape19.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape19.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape20 = sheet.Shapes.AddShape(81,sheet.cells(1,17).Left, sheet.cells(4,1).Top, 30, 30)
        shape20.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape20.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape21 = sheet.Shapes.AddShape(82,sheet.cells(1,19).Left, sheet.cells(4,1).Top, 30, 30)
        shape21.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape21.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape22 = sheet.Shapes.AddShape(73,sheet.cells(1,18).Left, sheet.cells(22,1).Top, 30, 30)
        shape22.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape22.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape22.TextFrame2.TextRange.Characters.Text = "F"
        shape22.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape22.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape23 = sheet.Shapes.AddShape(73,sheet.cells(1,19).Left, sheet.cells(22,1).Top, 30, 30)
        shape23.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape23.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape23.TextFrame2.TextRange.Characters.Text = "E"
        shape23.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape23.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape24 = sheet.Shapes.AddShape(73,sheet.cells(1,19).Left, sheet.cells(23,1).Top, 30, 30)
        shape24.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape24.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape24.TextFrame2.TextRange.Characters.Text = "E"
        shape24.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape24.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape25 = sheet.Shapes.AddShape(1,sheet.cells(1,19).Left, sheet.cells(20,1).Top, 30, 30)
        shape25.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape25.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape26 = sheet.Shapes.AddShape(73,sheet.cells(1,20).Left, sheet.cells(20,1).Top, 30, 30)
        shape26.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape26.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape26.TextFrame2.TextRange.Characters.Text = "〒"
        shape26.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape26.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape27 =sheet.Shapes.AddShape(73,sheet.cells(1,20).Left, sheet.cells(22,1).Top, 30, 30)
        shape27.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape27.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape27.TextFrame2.TextRange.Characters.Text = "〒"
        shape27.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape27.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape28 = sheet.Shapes.AddShape(81,sheet.cells(1,21).Left, sheet.cells(12,1).Top, 30, 30)
        shape28.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape28.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape29 = sheet.Shapes.AddShape(82,sheet.cells(1,22).Left, sheet.cells(13,1).Top, 30, 30)
        shape29.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape29.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape30 = sheet.Shapes.AddShape(4,sheet.cells(1,23).Left, sheet.cells(13,1).Top, 30, 30)
        shape30.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape30.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape31 = sheet.Shapes.AddShape(4,sheet.cells(1,22).Left, sheet.cells(12,1).Top, 30, 30)
        shape31.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape31.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape32 = sheet.Shapes.AddShape(4,sheet.cells(1,20).Left, sheet.cells(11,1).Top, 30, 30)
        shape32.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape32.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape33 = sheet.Shapes.AddShape(73,sheet.cells(1,23).Left, sheet.cells(12,1).Top, 30, 30)
        shape33.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape33.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape33.TextFrame2.TextRange.Characters.Text = "F"
        shape33.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape33.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape34 = sheet.Shapes.AddShape(82,sheet.cells(1,25).Left, sheet.cells(13,1).Top, 30, 30)
        shape34.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape34.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape35 = sheet.Shapes.AddShape(82,sheet.cells(1,25).Left, sheet.cells(14,1).Top, 30, 30)
        shape35.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape35.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape36 =sheet.Shapes.AddShape(77,sheet.cells(1,26).Left, sheet.cells(14,1).Top, 30, 30)
        shape36.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape36.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape37 =sheet.Shapes.AddShape(77,sheet.cells(1,26).Left, sheet.cells(12,1).Top, 30, 30)
        shape37.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape37.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape38 = sheet.Shapes.AddShape(73,sheet.cells(1,26).Left, sheet.cells(10,1).Top, 30, 30)
        shape38.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape38.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape38.TextFrame2.TextRange.Characters.Text = "F"
        shape38.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape38.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape39 = sheet.Shapes.AddShape(73,sheet.cells(1,27).Left, sheet.cells(10,1).Top, 30, 30)
        shape39.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape39.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape39.TextFrame2.TextRange.Characters.Text = "P"
        shape39.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape39.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape40 = sheet.Shapes.AddShape(81,sheet.cells(1,27).Left, sheet.cells(37,1).Top, 30, 30)
        shape40.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape40.Line.ForeColor.RGB = rgbToInt((0,0,0))
        
        shape41 = sheet.Shapes.AddShape(73,sheet.cells(1,28).Left, sheet.cells(35,1).Top, 30, 30)
        shape41.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape41.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape41.TextFrame2.TextRange.Characters.Text = "PC"
        shape41.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape41.TextFrame2.TextRange.Characters.Font.Size = 8
        shape41.TextFrame2.TextRange.Characters.Font.Spacing = -1 
        
        shape42 = sheet.Shapes.AddShape(73,sheet.cells(1,29).Left, sheet.cells(35,1).Top, 30, 30)
        shape42.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape42.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape42.TextFrame2.TextRange.Characters.Text = "K"
        shape42.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape42.TextFrame2.TextRange.Characters.Font.Size = 8
        
        shape43 = sheet.Shapes.AddShape(73,sheet.cells(1,32).Left, sheet.cells(35,1).Top, 30, 30)
        shape43.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape43.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape43.TextFrame2.TextRange.Characters.Text = "PC"
        shape43.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape43.TextFrame2.TextRange.Characters.Font.Size = 8
        shape43.TextFrame2.TextRange.Characters.Font.Spacing = -1 
        
        shape44 = sheet.Shapes.AddShape(73,sheet.cells(1,33).Left, sheet.cells(35,1).Top, 30, 30)
        shape44.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape44.Line.ForeColor.RGB = rgbToInt((0,0,0))
        #shape1.TextFrame.autofit_text("F")
        shape44.TextFrame2.TextRange.Characters.Text = "K"
        shape44.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = rgbToInt((0,0,0))
        shape44.TextFrame2.TextRange.Characters.Font.Size = 8
        
        
        shape45 = sheet.Shapes.AddShape(82,sheet.cells(1,33).Left, sheet.cells(37,1).Top, 30, 30)
        shape45.Fill.ForeColor.RGB = rgbToInt((255,255,255))
        shape45.Line.ForeColor.RGB = rgbToInt((0,0,0))
        shape45.TextFrame2.TextRange.Characters.Font.Size = 8
        
      
        #コネクターの書き方　
        connector1 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
#        connector1.BeginConnect.ConnectedShape = shape1
#        connector1.EndConnect.ConnectedShape = shape2
#        connector1.BeginConnect(ConnectedShape=shape1, ConnectionSite=1)
#        connector1.EndConnect (ConnectedShape=shape2, ConnectionSite=1)
        connector1.ConnectorFormat.BeginConnect(shape1, 5)
        connector1.ConnectorFormat.EndConnect(shape2, 1)
        
        connector2 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector2.ConnectorFormat.BeginConnect(shape2, 7)
        connector2.ConnectorFormat.EndConnect(shape3, 3)
        
        connector3 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector3.ConnectorFormat.BeginConnect(shape3, 7)
        connector3.ConnectorFormat.EndConnect(shape4, 3)
        
        connector4 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector4.ConnectorFormat.BeginConnect(shape4, 7)
        connector4.ConnectorFormat.EndConnect(shape5, 3)
        
        connector5 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector5.ConnectorFormat.BeginConnect(shape5, 1)
        connector5.ConnectorFormat.EndConnect(shape6, 3)
        
        connector6 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector6.ConnectorFormat.BeginConnect(shape6, 4)
        connector6.ConnectorFormat.EndConnect(shape7, 2)
        
        connector7 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector7.ConnectorFormat.BeginConnect(shape7, 4)
        connector7.ConnectorFormat.EndConnect(shape10, 3)
        
        connector8 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector8.ConnectorFormat.BeginConnect(shape8, 7)
        connector8.ConnectorFormat.EndConnect(shape9, 3)
        
        connector9 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector9.ConnectorFormat.BeginConnect(shape9, 7)
        connector9.ConnectorFormat.EndConnect(shape11, 3)
        
        connector10 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector10.ConnectorFormat.BeginConnect(shape12, 7)
        connector10.ConnectorFormat.EndConnect(shape13, 3)
        
        connector11 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector11.ConnectorFormat.BeginConnect(shape13, 7)
        connector11.ConnectorFormat.EndConnect(shape17, 2)
        
        connector12 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector12.ConnectorFormat.BeginConnect(shape14, 7)
        connector12.ConnectorFormat.EndConnect(shape15, 3)
        
        connector13 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector13.ConnectorFormat.BeginConnect(shape15, 7)
        connector13.ConnectorFormat.EndConnect(shape18, 3)
        
        connector14 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector14.ConnectorFormat.BeginConnect(shape16, 7)
        connector14.ConnectorFormat.EndConnect(shape20, 2)
        
        connector15 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector15.ConnectorFormat.BeginConnect(shape20, 4)
        connector15.ConnectorFormat.EndConnect(shape21, 2)
        
        connector16 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector16.ConnectorFormat.BeginConnect(shape18, 5)
        connector16.ConnectorFormat.EndConnect(shape19, 1)
        
        connector17 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector17.ConnectorFormat.BeginConnect(shape19, 4)
        connector17.ConnectorFormat.EndConnect(shape24, 3)
        
        connector18 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector18.ConnectorFormat.BeginConnect(shape22, 7)
        connector18.ConnectorFormat.EndConnect(shape23, 3)
        
        connector19 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector19.ConnectorFormat.BeginConnect(shape23, 7)
        connector19.ConnectorFormat.EndConnect(shape27, 3)
        
        connector20 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector20.ConnectorFormat.BeginConnect(shape25, 4)
        connector20.ConnectorFormat.EndConnect(shape26, 3)
        
        connector21 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector21.ConnectorFormat.BeginConnect(shape26, 1)
        connector21.ConnectorFormat.EndConnect(shape32, 3)
        
        connector22 = sheet.Shapes.AddConnector(2,150, 150, 150, 150)
        connector22.ConnectorFormat.BeginConnect(shape27, 7)
        connector22.ConnectorFormat.EndConnect(shape28, 3)
        
        connector23 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector23.ConnectorFormat.BeginConnect(shape28, 4)
        connector23.ConnectorFormat.EndConnect(shape31, 2)
        
        connector24 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector24.ConnectorFormat.BeginConnect(shape31, 4)
        connector24.ConnectorFormat.EndConnect(shape33, 3)
        
        connector25 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector25.ConnectorFormat.BeginConnect(shape33, 7)
        connector25.ConnectorFormat.EndConnect(shape37, 3)
        
        connector26 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector26.ConnectorFormat.BeginConnect(shape29, 4)
        connector26.ConnectorFormat.EndConnect(shape30, 2)
        
        connector27 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector27.ConnectorFormat.BeginConnect(shape30, 4)
        connector27.ConnectorFormat.EndConnect(shape34, 2)
        
        connector28 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector28.ConnectorFormat.BeginConnect(shape35, 4)
        connector28.ConnectorFormat.EndConnect(shape36, 3)
        
        connector29 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector29.ConnectorFormat.BeginConnect(shape38, 7)
        connector29.ConnectorFormat.EndConnect(shape39, 3)
        
        connector30 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector30.ConnectorFormat.BeginConnect(shape39, 5)
        connector30.ConnectorFormat.EndConnect(shape40, 1)
        
        connector31 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector31.ConnectorFormat.BeginConnect(shape40, 4)
        connector31.ConnectorFormat.EndConnect(shape45, 2)
        
        connector32 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector32.ConnectorFormat.BeginConnect(shape41, 7)
        connector32.ConnectorFormat.EndConnect(shape42, 3)
        
        connector33 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector33.ConnectorFormat.BeginConnect(shape42, 7)
        connector33.ConnectorFormat.EndConnect(shape43, 3)
        
        connector34 = sheet.Shapes.AddConnector(1,150, 150, 150, 150)
        connector34.ConnectorFormat.BeginConnect(shape43, 7)
        connector34.ConnectorFormat.EndConnect(shape44, 3)
        
        
        
        
    

        
        workbook.SaveAs(output_save_path) #saveas名前をつけて保存
        workbook.Close()
        excel.Quit()
        sheet = None
        workbook = None
        excel = None

        
#        reader = csv.reader(csv_file)
#        
#        for csv_row in reader:
#            print(csv_row)
       # csv_file.save(os.path.join('C:\\Users\\pi199\\Anaconda3\\project\\workflow\\workflow\\accounts\\input\\', csv_file.filename))
    return render(req, 'top.html', login_data)

@csrf_exempt
def csv_delete(req):
    input_path = os.path.join(input_dir, "input.csv")
    print(input_path)
    os.remove(input_path)
    return render(req, 'top.html', login_data)

