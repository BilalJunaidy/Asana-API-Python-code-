#Import all dependencies
import openpyxl
import win32com.client as win32
from datetime import datetime
import os
import asana
from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd


def get_workspaces():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True

    for wb in excel.Workbooks:
        if 'Client_Interface_To_Asana_API' in wb.Name:
            PAT_Sheet = wb.ActiveSheet
            PAT = PAT_Sheet.Range("B2").Value
            client = asana.Client.access_token(PAT)
            response = client.workspaces.get_workspaces(opt_fields=['gid','name'], opt_pretty=True)

            workspace_list = []
            i = 11
            for item in response:
                list = []
                list.append(item['name'])
                list.append(item['gid'])
                workspace_list.append(list)

            row = len(workspace_list)
            col = len(workspace_list[0])
            
            PAT_Sheet.Range(PAT_Sheet.Cells(8,1), PAT_Sheet.Cells(8+row-1,1+col-1)).Value = workspace_list
            PAT_Sheet.Columns.AutoFit()

            ##As part of the VBA, make sure to format the workspace id as a number and not as a general
            
#get_workspaces()


def get_projects():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    for wb in excel.Workbooks:
        if 'Client_Interface_To_Asana_API' in wb.Name:
            
            PAT_Sheet = wb.ActiveSheet
            workspace_gid = str(PAT_Sheet.Range("B8").Value)
            workspace_gid = str(workspace_gid.split('.')[0])
            PAT = PAT_Sheet.Range("B2").Value
            client = asana.Client.access_token(PAT)
            ##response = client.projects.get_projects({'organization': f'{workspace_gid}'}, opt_pretty=True)
            response = client.projects.get_projects_for_workspace(workspace_gid, opt_pretty=True)

            Project_sheet = wb.Worksheets.Add(After=wb.ActiveSheet)
            Project_sheet.Name = "Project Name and ID"
            Project_sheet = wb.Sheets(3)
            Project_sheet.Range("A1").Value = "Project name"
            Project_sheet.Range("B1").Value = "Project id"
            Project_sheet.Columns.AutoFit()

            project_list = []

            for item in response:
                list = []
                list.append(item['name'])
                list.append(item['gid'])
                project_list.append(list)

            row = len(project_list)
            col = len(project_list[0])
            
            Project_sheet.Range(Project_sheet.Cells(2,1), Project_sheet.Cells(2+row-1,1+col-1)).Value = project_list
            Project_sheet.Columns.AutoFit()
            
#get_projects()


def get_task():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    for wb in excel.Workbooks:
        if 'Client_Interface_To_Asana_API' in wb.Name:
            PAT_Sheet = wb.Sheets(1)
            PAT = PAT_Sheet.Range("B2").Value
            client = asana.Client.access_token(PAT)

            Project_sheet = wb.Sheets(2)
            project_gid = str(Project_sheet.Range("B3").Value)
            project_gid = project_gid
            #Tasks_sheet = wb.Worksheets.Add(After=wb.ActiveSheet)
            #Tasks_sheet.Name = "Task Name and ID"
            Tasks_sheet = wb.Sheets(3)
            Tasks_sheet.Range("A1").Value = "Task name"
            Tasks_sheet.Range("B1").Value = "Task id"

            response = client.tasks.get_tasks_for_project(project_gid, {'limit':'10'}, opt_fields=['gid','name'], opt_pretty=True)

            tasks_list = []
            for item in response:
                list = []
                tasks_name = item['name']
                list.append(tasks_name)
                tasks_gid = item['gid']
                list.append(tasks_gid)
                tasks_list.append(list)
                
                row = len(tasks_list)
                col = len(tasks_list[0])
            
                Tasks_sheet.Range(Tasks_sheet.Cells(2,1), Tasks_sheet.Cells(2+row-1,1+col-1)).Value = tasks_list
            Tasks_sheet.Columns("B").AutoFit()


##get_task()


def get_task_details():
    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    for wb in excel.Workbooks:
        if 'Client_Interface_To_Asana_API' in wb.Name:
            PAT_Sheet = wb.Sheets(1)
            PAT = PAT_Sheet.Range("B2").Value
            client = asana.Client.access_token(PAT)

            tasks_gid = str(wb.Sheets(3).Range("B22").Value)
            tasks_name = str(wb.Sheets(3).Range("A22").Value)
            tasks_list = []
            tasks_list.append(f'{tasks_gid}')
            
            sub_tasks = client.tasks.get_subtasks_for_task(tasks_gid, {'limit':'10'}, opt_fields=['gid','name'], opt_pretty=True)
            for item in sub_tasks:
                tasks_list.append(item['gid'])

            output_list = []
            
            for i, tasks_gid in enumerate(tasks_list):
                response = client.tasks.get_task(tasks_gid, opt_pretty=True, opt_fields = ['name','created_at','completed','assignee', 'modified_at', 'due_on','start_on','tags','notes','parent', 'section'])

                if i > 0:
                    response['parent'] = tasks_name

                if len(response['tags']) != 0:
                    response['tags'] = response['tags'][0]['name']
                
                if not response['assignee'] is None:    
                    user_gid = response['assignee']['gid']
                    result = client.users.get_user(user_gid,opt_pretty=True, opt_fields=['name','email'])
                    response['assignee'] = result['name']
                    response['assignee email'] = result['email']
            
                output_list.append(response)

            df = pd.DataFrame(output_list)

            df['created_at'] = df['created_at'].str.split("T", expand=True)
            df['modified_at'] = df['modified_at'].str.split("T", expand=True)
            
            df.to_csv('ClientInterfaceOutput.csv', index=False)

get_task_details()




































