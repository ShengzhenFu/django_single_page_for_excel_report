from django.shortcuts import render
from .backend import main_app
from .backend import incident_status_optimized
import time
from django.template import loader


def file_read(fname):
    with open(fname) as f:
        content_list = f.readlines()
        max_line = len(content_list)
        # print(content_list[max_line-10:max_line])
        return content_list[max_line-16:max_line]


def button(request):
    return render(request, 'index.html')


def output(request):
    run = main_app.main()
    time.sleep(5)
    filepath = 'D:\\projects\\pmo_reports\\pmo_report.log'
    run_log = file_read(filepath)
    run_log_str = ""
    for x in run_log:
        run_log_str += x + '\r\n'
    return render(request, 'index.html', context={'data': run + '\r\n' + run_log_str})


def index(request):
    gen_month = 'May 2019'
    context ={
        'gen_month': gen_month,
        }
    template = loader.get_template('index.html')
    return render(request, 'index.html', context)

