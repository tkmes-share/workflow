#!/usr/bin/env python
# -*- coding: utf-8 -*-
from flask import Flask, request, session, g, redirect, url_for, abort, render_template, flash, jsonify
from datetime import datetime
from pytz import timezone
import MySQLdb
import codecs
import os
import sys
import csv

# global変数初期化
user_id=""
user_password=""
login_date=""
entries={}

main = Flask(__name__)

# /login にアクセスしたときの処理
@main.route("/")
def login():
   return render_template('login.html', entries=entries)

# ログインページから/top にアクセスしたときの処理（POST）
@main.route('/top_post', methods=['POST'])
def top_post():
    global user_id
    global user_password
    global login_date
    global entries

    # TB_CORPから全件取得し、modelにセット
    connection = MySQLdb.connect(db="c9", user="tkmes", passwd="tkmesadmin", charset="utf8")
    cursor=connection.cursor()
#    sql = "select a.USER_ID, a.CORP_NAME, a.CORP_PLACE, b.DEAL_PAYMENT, b.DATETIME from TB_CORP_REGISTER_DEMO as a left join TB_DEAL_HISTORY_DEMO as b where DATETIME = (SELECT MAX(DATETIME) FROM TB_DEAL_HISTORY_DEMO as bs WHERE b.USER_ID=bs.USER_ID and b.CORP_NAME=bs.CORP_NAME) on a.USER_ID = b.USER_ID and a.CORP_NAME = b.CORP_NAME"
    sql = "select * from TB_DEAL_HISTORY_DEMO as b where DATETIME = (SELECT MAX(DATETIME) FROM TB_DEAL_HISTORY_DEMO as bs WHERE b.USER_ID=bs.USER_ID and b.CORP_NAME=bs.CORP_NAME)"
    cursor.execute(sql)
    entries = cursor.fetchall()
    cursor.close()
    connection.close()

    if request.method == 'POST':
       user_id = request.form['id']
       user_password = request.form['password']
       login_date = datetime.now(timezone('Asia/Tokyo')).strftime("%Y/%m/%d %H:%M")
#       print (user_id, user_password, login_date, entries)
       return render_template('top.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)
    else:
        return redirect(url_for('login'))

# /top にアクセスしたときの処理（GET）
@main.route('/top', methods=['GET'])
def top():
    global entries
    # TB_CORPから全件取得し、modelにセット
    connection = MySQLdb.connect(db="c9", user="tkmes", passwd="tkmesadmin", charset="utf8")
    cursor = connection.cursor()
    sql = "select * from TB_DEAL_HISTORY_DEMO as b where DATETIME = (SELECT MAX(DATETIME) FROM TB_DEAL_HISTORY_DEMO as bs WHERE b.USER_ID=bs.USER_ID and b.CORP_NAME=bs.CORP_NAME)"
    cursor.execute(sql)
    entries = cursor.fetchall()
    cursor.close()
    connection.close()

    return render_template('top.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/privacyPolicy")
def privacyPolicy():
   return render_template('privacyPolicy.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/newRegister")
def newRegister():
   return render_template('newRegister.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/newRegisterNext")
def newRegisterNext():
   return render_template('newRegisterNext.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/newRegisterLast")
def newRegisterLst():
   return render_template('newRegisterLast.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/newRegisterComplete")
def newRegisterComplete():
   return render_template('newRegisterComplete.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/passChange")
def passChange():
   return render_template('passChange.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)
        
@main.route("/passChangeComplete")
def passChangeComplete():
   return render_template('passChangeComplete.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/help")
def helpQA():
   return render_template('help.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/scoring")
def scoring():
   return render_template('scoring.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/scoring2")
def scoring2():
   return render_template('scoring2.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/scoring3")
def scoring3():
   return render_template('scoring3.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)   

@main.route("/scoring4")
def scoring4():
   return render_template('scoring4.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)   

@main.route("/scoring5")
def scoring5():
   return render_template('scoring5.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)   

@main.route("/dbEdit")
def dbEdit():
    # TBから全件取得し、modelにセット
    connection = MySQLdb.connect(db="c9", user="tkmes", passwd="tkmesadmin", charset="utf8")
    cursor = connection.cursor()
    sql = "select * from TB_DEAL_HISTORY_DEMO as b where CORP_ID = '0001' and DATETIME = (SELECT MAX(DATETIME) FROM TB_DEAL_HISTORY_DEMO as bs WHERE b.USER_ID=bs.USER_ID and b.CORP_NAME=bs.CORP_NAME)"
    cursor.execute(sql)
    entries = cursor.fetchall()
    cursor.close()
    connection.close()
    print(entries)

    return render_template('dbEdit.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries) 

@main.route("/dbEdit2")
def dbEdit2():
    # TBから全件取得し、modelにセット
    connection = MySQLdb.connect(db="c9", user="tkmes", passwd="tkmesadmin", charset="utf8")
    cursor = connection.cursor()
    sql = "select * from TB_DEAL_HISTORY_DEMO as b where CORP_ID = '0002' and DATETIME = (SELECT MAX(DATETIME) FROM TB_DEAL_HISTORY_DEMO as bs WHERE b.USER_ID=bs.USER_ID and b.CORP_NAME=bs.CORP_NAME)"
    cursor.execute(sql)
    entries = cursor.fetchall()
    cursor.close()
    connection.close()
    print(entries)

    return render_template('dbEdit2.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries) 

@main.route("/dbEdit3")
def dbEdit3():
    # TBから全件取得し、modelにセット
    connection = MySQLdb.connect(db="c9", user="tkmes", passwd="tkmesadmin", charset="utf8")
    cursor = connection.cursor()
    sql = "select * from TB_DEAL_HISTORY_DEMO as b where CORP_ID = '1766' and DATETIME = (SELECT MAX(DATETIME) FROM TB_DEAL_HISTORY_DEMO as bs WHERE b.USER_ID=bs.USER_ID and b.CORP_NAME=bs.CORP_NAME)"
    cursor.execute(sql)
    entries = cursor.fetchall()
    cursor.close()
    connection.close()
    print(entries)

    return render_template('dbEdit3.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries) 

@main.route("/dbEdit4")
def dbEdit4():
    # TBから全件取得し、modelにセット
    connection = MySQLdb.connect(db="c9", user="tkmes", passwd="tkmesadmin", charset="utf8")
    cursor = connection.cursor()
    sql = "select * from TB_DEAL_HISTORY_DEMO as b where CORP_ID = '1801' and DATETIME = (SELECT MAX(DATETIME) FROM TB_DEAL_HISTORY_DEMO as bs WHERE b.USER_ID=bs.USER_ID and b.CORP_NAME=bs.CORP_NAME)"
    cursor.execute(sql)
    entries = cursor.fetchall()
    cursor.close()
    connection.close()
    print(entries)

    return render_template('dbEdit4.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries) 

@main.route("/dbEdit5")
def dbEdit5():
    # TBから全件取得し、modelにセット
    connection = MySQLdb.connect(db="c9", user="tkmes", passwd="tkmesadmin", charset="utf8")
    cursor = connection.cursor()
    sql = "select * from TB_DEAL_HISTORY_DEMO as b where CORP_ID = '7875' and DATETIME = (SELECT MAX(DATETIME) FROM TB_DEAL_HISTORY_DEMO as bs WHERE b.USER_ID=bs.USER_ID and b.CORP_NAME=bs.CORP_NAME)"
    cursor.execute(sql)
    entries = cursor.fetchall()
    cursor.close()
    connection.close()
    print(entries)

    return render_template('dbEdit5.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries) 

@main.route("/registerRequest")
def registerRequest():
   return render_template('registerRequest.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/registerRequestComplete")
def registerRequestComplete():
   return render_template('registerRequestComplete.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/registerEdit")
def registerEdit():
   return render_template('registerEdit.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/registerEditComplete")
def registerEditComplete():
   return render_template('registerEditComplete.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/userRegister")
def userRegister():
   return render_template('userRegister.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

@main.route("/userRegisterComplete")
def userRegisterComplete():
   return render_template('userRegisterComplete.html', user_id=user_id, user_password=user_password, login_date=login_date, entries=entries)

main.config['UPLOADED_PATH'] = os.getcwd() + '/upload'

@main.route('/csv_upload', methods=['POST', 'GET'])
def csv_upload():
    if request.method == 'POST':
        for f in request.files.getlist('file'):
               f.save(os.path.join(main.config['UPLOADED_PATH'], f.filename))
    return redirect(url_for('csv_db_insert'))

@main.route('/csv_delete', methods=['POST', 'GET'])
def csv_delete():
    if request.headers['Content-Type'] != 'application/json':
        print(request.headers['Content-Type'])
        return jsonify(_return='error'), 400

    f = request.json
    os.remove(os.path.join(main.config['UPLOADED_PATH'], f))
    return jsonify(_return='delete ok')

@main.route('/csv_db_insert', methods=['POST', 'GET'])
def csv_db_insert():
  f = open("./upload/取引履歴.csv", "r")
  reader = csv.reader(f)
  header = next(reader)

  conn = MySQLdb.connect(db="c9", user="tkmes", passwd="tkmesadmin", charset="utf8")
  cursor=conn.cursor()
  for row in reader:
    sql = "INSERT INTO TB_DEAL_HISTORY_DEMO values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
    cursor.execute(sql, ("tkmes",row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13],row[14],datetime.now(timezone('Asia/Tokyo')).strftime("%Y/%m/%d %H:%M:%S")))
    conn.commit()
  cursor.close()
  conn.close()
  f.close()

  return jsonify(_return='upload ok')

if __name__ == "__main__":
    main.run(host='0.0.0.0', port=8080)
