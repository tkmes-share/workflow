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
    if req.method == 'POST':
        post_data = io.TextIOWrapper(req.FILES['file'])
        csv_file = csv.reader(post_data)
        header = next(csv_file)

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
        sheet.Cells(16,1).Value="仕入れ先"
        sheet.Cells(31,1).Value="経理部"


        
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
    return render(req, 'top.html', login_date)

@csrf_exempt
def csv_delete(req):
    input_path = os.path.join(input_dir, "input.csv")
    print(input_path)
    os.remove(input_path)
    return render(req, 'top.html', login_data)

