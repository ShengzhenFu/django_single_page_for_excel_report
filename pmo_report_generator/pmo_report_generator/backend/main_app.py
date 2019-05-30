import logging
import logging.handlers as handlers
import sys
from .openactionitem import sheet3_update
from .root_cause_analysis import sheet4_update
from .responseTime import sheet2_update
from shutil import copyfile
import os
import configparser
from .copy_raw_data import sheet567_update
from .incident_status_optimized import sheet1_update

"""""""""""""""""""""""""""
author: Shengzhen.Fu
purpose: auto generate PMO report based on the source of excels
version: 1.0.0
date: 2/Feb/2019
----------------------------------------------------------------
updated by: Shengzhen.Fu
update content: download source file from shared folder before update action
version: 1.0.1
date: 13/Feb/2019
----------------------------------------------------------------
"""""""""""""""""""""""""""
# log initialization
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
log_path_config = 'D:\\projects\\pmo_reports\\pmo_report.log'
logHandler = handlers.RotatingFileHandler(log_path_config, maxBytes=20000000, backupCount=10)
logHandler.setLevel(logging.INFO)
logHandler.setFormatter(formatter)

logger.addHandler(logHandler)

# retrieve global configs and log_config
config_path = "D:\\projects\\pmo_reports\\"
config = configparser.ConfigParser()
config.read(config_path + 'config.ini')

# define the source and destination files
# src_file_p1p2 = '\\\SZ-G3YC6H2\\Users\\dluo\\OneDrive - MORNINGSTAR INC\\Incident Data\\Raw Data\\P1&P2 Incident 201701_201901_All.xlsx'
# src_file_action = '\\\SZ-G3YC6H2\\Users\\dluo\\OneDrive - MORNINGSTAR INC\\Incident Data\\Raw Data\\Action Items List 201801_2018012.xlsx'
# dst_file_p1p2 = 'D:\\projects\\pmo_reports\\P1&P2 Incident 201701_201901_All.xlsx'
# dst_file_action = 'D:\\projects\\pmo_reports\\Action Items List 201801_2018012.xlsx'
# report_file = 'D:\\projects\\pmo_reports\\ReportTemplate_Kun_dev.xlsx'

src_file_p1p2 = config.get('source', 'src_p1p2_file')
src_file_action = config.get('source', 'src_action_file')
dst_file_p1p2 = config.get('destination', 'dst_p1p2_file')
dst_file_action = config.get('destination', 'dst_action_file')
report_file = config.get('destination', 'dst_report_file')


def download_source_file():
    # download latest P1&P2 Incident 201701_201901_All.xlsx from shared folder
    if os.path.exists(src_file_p1p2):
        print('file', src_file_p1p2, 'exists, going to download it to local')
        src1 = src_file_p1p2
        dst1 = dst_file_p1p2
        try:
            copyfile(src1, dst1)
        except IOError as e:
            print('unable to download file', src1, 'error is', e)
        print('complete download', src_file_p1p2, 'to local', dst_file_p1p2)
    else:
        print('File Not Found, please verify file exist at', src_file_p1p2)

    # download latest Action Items List 201801_2018012.xlsx from shared folder
    if os.path.exists(src_file_action):
        print('file', src_file_action, 'exists, going to download it to local')
        src2 = src_file_action
        dst2 = dst_file_action
        try:
            copyfile(src2, dst2)
        except IOError as e:
            print('unable to download file', src2, 'error is', e)
            # print('start to copy via  system method instead of shutil.copyfile')
            # os.system('copy "%s" "%s"' % (src2, dst2))
        # print('complete download', src_file_action, 'to local', dst_file_action)
    else:
        print('File Not Found, please verify file exist at', src_file_action)


def main():
    # download latest source excel
    # download_source_file()
    # start to read and update pmo report
    try:
        logger.info('start to update the create & resolved incidents data')
        sheet1_update.update_created_resolved_incident(dst_file_action, report_file)
        logger.info('complete update the create & resolved incidents data')
    except Exception as e:
        print("error %s occurred when update created & resolved incidents data", e)
        logger.exception('error %s occurred when update created & resolved incidents data', e)
        sys.exit(2)
    try:
        logger.info('start to update the response time data')
        sheet2_update.update_response_time_sheet(dst_file_p1p2, report_file)
        logger.info('complete update the response time data')
    except Exception as e:
        print("error %s occurred when update response time ", e)
        logger.exception('error %s occurred when update the response time data', e)
        sys.exit(2)
    try:
        logger.info('start to update the  open actionItems data')
        sheet3_update.update_open_action_data_new(dst_file_action, report_file)
        logger.info('completed update the open actionItems data')
    except Exception as e:
        print("error %s occurred when update open actionItems", e)
        logger.exception('error %s occurred when update the open actionItems data', e)
        sys.exit(2)
    try:
        logger.info('start to update the rca category break down data')
        sheet4_update.update_rca_data(dst_file_p1p2, report_file)
        logger.info('completed update the rca category break down data')
    except Exception as e:
        print("error %s occurred when update rca category breakdown data", e)
        logger.exception('error %s occurred when update the rca category breakdown data', e)
        sys.exit(2)
    try:
        logger.info('start to copy data to open action item list')
        sheet567_update.copy_open_actionitemlist(dst_file_action, report_file)
        logger.info('complete copy data to open action item list')
    except Exception as e:
        print("error %s occurred when copy data to open action item list", e)
        logger.exception('error %s occurred when copy data to open action item list', e)
        sys.exit(2)
    try:
        logger.info('start to copy data to Incident Backlog')
        sheet567_update.update_incident_backlog(dst_file_action, dst_file_p1p2, report_file)
        logger.info('complete copy data to Incident Backlog')
    except Exception as e:
        print("error %s occurred when copy data to Incident Backlog", e)
        logger.exception('error %s occurred when copy data to Incident Backlog', e)
        sys.exit(2)
    try:
        logger.info('start to copy data to Incident list ')
        sheet567_update.update_incident_list(dst_file_p1p2, report_file)
        logger.info('complete copy data to Incident list ')
    except Exception as e:
        print("error %s occurred when copy data to Incident list", e)
        logger.exception('error %s occurred when copy data to Incident list', e)
        sys.exit(2)
    return 'successfully update report'


if __name__ == '__main__':
    main()
