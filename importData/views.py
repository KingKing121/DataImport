import datetime
import os

import cx_Oracle
import xlrd as xlrd
from django import forms
from django.http import HttpResponse
from django.shortcuts import render_to_response, render


def sayHello(request):
    s = 'Hello World!'
    current_time = datetime.datetime.now()
    html = '<html><head></head><body><h1> %s </h1><p> %s </p></body></html>' % (s, current_time)
    return HttpResponse(html)


def student(request):
    list = [{id: 1, 'name': 'Jack'}, {id: 2, 'name': 'Rose'}]
    return render_to_response('student.html', {'students': list})


def validate_excel(value):
    print(value.name)
    e_name = value.name
    # 这里可以加验证 搜一下怎么加验证跳转。
    zhui = e_name[e_name.index("."):]


class UploadExcelForm(forms.Form):
    excel = forms.FileField(validators=[validate_excel])


def uploadFile(request):
    # 解决乱码
    os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
    form = UploadExcelForm(request.POST, request.FILES)
    print(form.is_valid())
    if form.is_valid():
        wb = xlrd.open_workbook(filename=None, file_contents=request.FILES['excel'].read())  # 关键点在于这里
        table = wb.sheets()[0]
        # 上传文件名  可以根据文件名称控制 xls 或者 xlsx
        # 如果是xlsx 则换 workbook = load_workbook('D:\\Py_exercise\\test_openpyxl.xlsx')
        e_name = request.FILES['excel']
        print(e_name)

        dfun = []
        nrows = table.nrows  # 行数

        # 获取行数据,并删除合计行 从第三行开始获取
        for i in range(2, nrows):
            dfun.append(table.row_values(i))

        # 由于合计行位置不正确。调整数据。
        for index, item in enumerate(dfun):
            if item[0] == "合计":
                ss = dfun[index - 1]
                item[0] = ss[0] + 1  # 合计行加行号
                item[1] = ss[1]  # 合计行年份
                item[2] = ss[2]  # 合计行月份
                item[3] = "合计"
            if item[0] == "总合计":  # 同理
                ss = dfun[index - 2]
                item[0] = ss[0] + 2
                item[1] = ss[1]
                item[2] = ss[2]
                item[3] = "总合计"

        # 输出数据,看是否符合标准
        print(len(dfun))

        # 连接数据库
        username = "apps"
        userpwd = "apps"
        host = "20.20.1.20"
        port = 1521
        dbname = "PROD"
        dsn = cx_Oracle.makedsn(host, port, dbname)
        db = cx_Oracle.connect(username, userpwd, dsn)

        # 获取游标
        cur = db.cursor()

        for item in dfun:
            for index, i in enumerate(item):
                if index == 0:  # 第一列值取整
                    item[index] = int(i)
                if i == '' and 15 > index > 5:  # null值替换0.00
                    item[index] = 0.00
                if 15 > index > 5:  # 小数点位数控制
                    item[index] = "%.7f" % item[index]

        insert_sql = "insert into CUX_BI_electricity (seq_num, elec_year,elec_month,factory_name, return_circui,install_site,last_mon_elec_meter,cur_mon_elec_meter,elec_rate,elec_kwh,loss_kwh,elec_sum_kwh,elec_price,elec_fee,deduction_tax_fee,elec_notes)  values(:1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13, :14, :15, :16)"

        # for item in dfun:  # 输出看值是否正确
        # print(item)
        # 执行sql语句
        cur.executemany(insert_sql, dfun)
        # 数据库操作提交
        db.commit()

        return HttpResponse("耶 ，你完成啦！")

    else:
        return render(request, "upload.html", {})
