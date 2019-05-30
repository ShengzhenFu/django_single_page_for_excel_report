import os
import pandas as pd
from .incident_status_optimized import sheet1_update

class sheet567_update:

    def copy_open_actionitemlist(srcfile, tarfile):
        """
        Params: source file name, target file name
        source excel: Action Items List 201801_201905.xlsx
        source sheet: <Action Items List>

        target excel: Incident Report for Apr_as of Apr 30 updated.xlsx
        target sheet: Open Action Item List

        copy open action item (action items resolved is null) from source excel sheet to target excel sheet
        """
        if os.path.exists(srcfile):
            # column F action items ticket, H action items resolved I Duration L team
            pd.options.mode.chained_assignment = None  # avoid warnings when add column to pandas dataframe
            df = pd.read_excel(srcfile, sheet_name='Action Items List', usecols="A,B,C,D,E,F,G,H,J,K,L,M")
            df1 = df[(df['action items resolved']).isnull()]
            # df1['Open Duration'] = df1['Open Duration'].map(lambda  x:(['Report Date']-['action items created']))
            print(df1)
            sheet1_update.append_df_to_excel(tarfile, df1, sheet_name='Open Action Item List', header=None, index=False, startrow=1,
                               columns=['key', 'priority', 'summary', 'created', 'rca key', 'action items ticket',
                                        'action items created', 'action items resolved', 'Report Date', 'Open Duration', 'team', 'comments'])

    def get_open_incident_key(src_actionitem):
        if os.path.exists(src_actionitem):
            # column A key, B priority C summary H action items resolved I Duration L team
            df = pd.read_excel(src_actionitem, sheet_name='Action Items List', usecols="A,H")
            df1 = df[(df['action items resolved']).isnull()].groupby(['key'])['action items resolved'].max().reset_index(name='rca')
            df2 = pd.DataFrame(df1, columns=['key'])
            list_rca = list(df2.key)
            print('total ', len(list_rca), 'open action items found')
            return list_rca

    def update_incident_backlog(src_actionitem, src_incident, tarfile):
        """
        Params: source actionitem file name, source incident file name,target file name

        source excel 1: Action Items List 201801_201905.xlsx
        source sheet 1: <Action Items List>
        source excel 2: P1&P2 Incident 201701_201905_All.xlsx
        source sheet 2: <ALL>

        target excel: Incident Report for Apr_as of Apr 30 updated.xlsx
        target sheet: Open Action Item List

        Get the Key of open action item
            action items resolved is null) from source excel 1 sheet 1
        Use the key to mapping to source excel 2
        copy a sub-collection of data from source to target sheet
        """
        open_action_keys = sheet567_update.get_open_incident_key(src_actionitem)
        if os.path.exists(src_incident):
            # pay attention to usecols below, they are mapping between <p1p2 incident excel - ALL sheet> and <Incident Report excel - Open Action Item List>
            df = pd.read_excel(src_incident, sheet_name='ALL', usecols="A,B,C,D,E,F,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AE,AF,AG")
            df1 = df[(df['key']).isin(open_action_keys)].sort_values(by=['created'])
            print(df1)
            sheet1_update.append_df_to_excel(tarfile, df1, sheet_name='Incident Backlog', header=None, index=False, startrow=1)

    def update_incident_list(src_incident, tarfile):
        """
        Params: source file name, target file name

        source excel: P1&P2 Incident 201701_201905_All.xlsx
        source sheet: <ALL>

        target excel: Incident Report for Apr_as of Apr 30 updated.xlsx
        target sheet: Incident list as of Apr 30

        copy a sub-collection data from source which
        created > 2019-01-01 and
        sz involved = Yes
        to target sheet
        """
        if os.path.exists(src_incident):
            # pay attention to usecols below, they are mapping between <p1p2 incident excel - ALL sheet> and <Incident Report excel - Open Action Item List>
            df = pd.read_excel(src_incident, sheet_name='ALL', usecols="A,B,C,D,E,F,G,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AE,AF,AG")
            df1 = df[(df['created'] > '2019-01-01') & (df['sz involved'] == 'Yes')].sort_values(by=['created'])
            print(df1)
            sheet1_update.append_df_to_excel(tarfile, df1, sheet_name='Incident list as of Apr 30', header=None, index=False, startrow=1)


