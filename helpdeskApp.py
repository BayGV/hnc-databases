import pyodbc as db
import pandas as pd
from tkinter import ttk
from tkinter import *
from tkinter import messagebox

conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')

#start initialisation of tkinter gui elements
root = Tk()
root.title('IT Helpdesk')
root.geometry('1000x800')

notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

#
#Call log page content
#
callLogPage = Frame(notebook)
notebook.add(callLogPage, text='Call Log')
l = Label(callLogPage, text='Call Log Information')
l.config(font=('Arial', 12))
l.pack(pady=10)

#configure treeview style
style = ttk.Style()
style.theme_use('default')
style.configure('Treeview',
                background='lightgray',
                foreground='black',
                rowheight=20,
                fieldbackground='lightgray')
style.map('Treeview', background=[('selected',
                                   'darkblue')])

#set up the treeview frame
callLogTreeFrame = Frame(callLogPage)
callLogTreeFrame.pack(padx=5, pady=10, anchor='n')
callLogTreeScrollY = Scrollbar(callLogTreeFrame, orient='vertical')
callLogTreeScrollY.pack(side=RIGHT, fill=Y)
callLogScrollX = Scrollbar(callLogTreeFrame, orient='horizontal')
callLogScrollX.pack(side=BOTTOM, fill=X)

callLogTree = ttk.Treeview(callLogTreeFrame,
                           yscrollcommand=callLogTreeScrollY.set,
                           xscrollcommand=callLogScrollX.set,
                           selectmode='extended')
callLogTree.pack()

#set up the treeview content
callLogTree['columns'] = ('issue_id',
                          'employee_name',
                          'ho.operator_name',
                          'call_time',
                          'call_date',
                          'asset_type',
                          'asset_make',
                          'operating_system',
                          'software_name',
                          'valid_license?',
                          'issue_type',
                          'issue_description',
                          'ho2.operator_name',
                          'issue_closed?',
                          'closed_time',
                          'closed_date',
                          'resolution_description',
                          'minutes_taken_to_resolve')
callLogTree.column('#0', width=0, stretch=False)
callLogTree.column('issue_id', anchor=CENTER, stretch=False, width=50)
callLogTree.column('employee_name', anchor=CENTER, stretch=False, width=120)
callLogTree.column('ho.operator_name', anchor=CENTER, stretch=False, width=120)
callLogTree.column('call_time', anchor=CENTER, stretch=False, width=100)
callLogTree.column('call_date', anchor=CENTER, stretch=False, width=75)
callLogTree.column('asset_type', anchor=CENTER, stretch=False, width=75)
callLogTree.column('asset_make', anchor=CENTER, stretch=False, width=75)
callLogTree.column('operating_system', anchor=CENTER, stretch=False, width=75)
callLogTree.column('software_name', anchor=CENTER, stretch=False, width=120)
callLogTree.column('valid_license?', anchor=CENTER, stretch=False, width=100)
callLogTree.column('issue_type', anchor=CENTER, stretch=False, width=120)
callLogTree.column('issue_description', anchor=CENTER, stretch=False, width=150)
callLogTree.column('ho2.operator_name', anchor=CENTER, stretch=False, width=120)
callLogTree.column('issue_closed?', anchor=CENTER, stretch=False, width=75)
callLogTree.column('closed_time', anchor=CENTER, stretch=False, width=100)
callLogTree.column('closed_date', anchor=CENTER, stretch=False, width=75)
callLogTree.column('resolution_description', anchor=CENTER, stretch=False, width=150)
callLogTree.column('minutes_taken_to_resolve', anchor=CENTER, stretch=False, width=100)

callLogTree.heading('#0', text='', anchor=CENTER)
callLogTree.heading('issue_id', text='Issue ID' ,anchor=CENTER)
callLogTree.heading('employee_name', text='Employee Name', anchor=CENTER)
callLogTree.heading('ho.operator_name', text='Reporting Operator', anchor=CENTER)
callLogTree.heading('call_time', text='Call Time', anchor=CENTER)
callLogTree.heading('call_date', text='Call Date' ,anchor=CENTER)
callLogTree.heading('asset_type', text='Asset Type', anchor=CENTER)
callLogTree.heading('asset_make', text='Asset Make', anchor=CENTER)
callLogTree.heading('operating_system', text='OS', anchor=CENTER)
callLogTree.heading('software_name', text='Software Name', anchor=CENTER)
callLogTree.heading('valid_license?', text='Valid License?' ,anchor=CENTER)
callLogTree.heading('issue_type', text='Issue Type', anchor=CENTER)
callLogTree.heading('issue_description', text='Issue Description', anchor=CENTER)
callLogTree.heading('ho2.operator_name', text='Assigned Operator', anchor=CENTER)
callLogTree.heading('issue_closed?', text='Issue Closed?', anchor=CENTER)
callLogTree.heading('closed_time', text='Closed Time', anchor=CENTER)
callLogTree.heading('closed_date', text='Closed Date', anchor=CENTER)
callLogTree.heading('resolution_description', text='Resolution Description', anchor=CENTER)
callLogTree.heading('minutes_taken_to_resolve', text='Minutes Taken To Resolve', anchor=CENTER)

callLogTreeLabelFrame = LabelFrame(callLogPage, text='View call log details here.')
callLogTreeLabelFrame.pack(fill='x', expand=False, padx=10, pady=2)

#call log search frame
searchCallLogFrame = LabelFrame(callLogPage,
                                text='Search for a specific call here.')
searchCallLogFrame.pack(fill='x', expand=False, padx=10, pady=2)

searchIssueIdCallLogLabel = Label(searchCallLogFrame, text='Issue ID:')
searchIssueIdCallLogLabel.grid(row=0, column=0, padx=10, pady=10)

searchIssueIdCallLogEntry = Entry(searchCallLogFrame)
searchIssueIdCallLogEntry.grid(row=0, column=1, padx=10, pady=10)

searchEmployeeNameCallLogLabel = Label(searchCallLogFrame, text='Employee name:')
searchEmployeeNameCallLogLabel.grid(row=0, column=4, padx=10, pady=10)

searchEmployeeNameCallLogEntry = Entry(searchCallLogFrame)
searchEmployeeNameCallLogEntry.grid(row=0, column=5, padx=10, pady=10)

searchIssueTypeCallLogLabel = Label(searchCallLogFrame, text='Issue type:')
searchIssueTypeCallLogLabel.grid(row=0, column=8, padx=10, pady=10)

issueType = ['Password reset', 'Software update', 'MFA', 'Hardware', 'Networking',
             'Permissions']
issueTypeVariable = StringVar(searchCallLogFrame)

searchIssueTypeCallLogEntry = OptionMenu(searchCallLogFrame, issueTypeVariable,
                                         *issueType)
searchIssueTypeCallLogEntry.grid(row=0, column=9, padx=10, pady=10)

searchAssetMakeCallLogLabel = Label(searchCallLogFrame, text='Asset make:')
searchAssetMakeCallLogLabel.grid(row=0, column=12, padx=10, pady=10)

searchAssetMakeCallLogEntry = Entry(searchCallLogFrame)
searchAssetMakeCallLogEntry.grid(row=0, column=13, padx=10, pady=10)

#operator call log search frame
searchOperatorCallLogFrame = LabelFrame(callLogPage,
                                text='Search for open operator calls here.')
searchOperatorCallLogFrame.pack(fill='x', expand=False, padx=10, pady=2)

searchOperatorsCallLogLabel = Label(searchOperatorCallLogFrame, text='Operator name:')
searchOperatorsCallLogLabel.grid(row=0, column=0, padx=10, pady=10)

searchOperatorsCallLogEntry = Entry(searchOperatorCallLogFrame)
searchOperatorsCallLogEntry.grid(row=0, column=1, padx=10, pady=10)

#update call log frame
updateCallLogFrame = LabelFrame(callLogPage,
                                text='Update an existing call here.')
updateCallLogFrame.pack(fill='x', expand=False, padx=10, pady=2)

issueToUpdateCallLogLabel = Label(updateCallLogFrame, text='Issue Id to update*:')
issueToUpdateCallLogLabel.grid(row=0, column=0, padx=10, pady=10)

issueToUpdateCallLogEntry = Entry(updateCallLogFrame)
issueToUpdateCallLogEntry.grid(row=0, column=1, padx=10, pady=10)

updateIssueTypeCallLogLabel = Label(updateCallLogFrame, text='Issue type*:')
updateIssueTypeCallLogLabel.grid(row=0, column=3, padx=10, pady=10)

updateIssueTypeCallLogEntry = OptionMenu(updateCallLogFrame, issueTypeVariable,
                                         *issueType)
updateIssueTypeCallLogEntry.grid(row=0, column=4, padx=10, pady=10)

updateIssueClosedCallLogLabel = Label(updateCallLogFrame, text='Issue closed?*:')
updateIssueClosedCallLogLabel.grid(row=0, column=6, padx=10, pady=10)

updateIssueClosedCallLogEntry = Entry(updateCallLogFrame)
updateIssueClosedCallLogEntry.grid(row=0, column=7, padx=10, pady=10)

updateResolveDateCallLogLabel = Label(updateCallLogFrame, text='Resolve date (yyyy-mm-dd):')
updateResolveDateCallLogLabel.grid(row=0, column=9, padx=10, pady=10)

updateResolveDateCallLogEntry = Entry(updateCallLogFrame)
updateResolveDateCallLogEntry.grid(row=0, column=10, padx=10, pady=10)

updateResolveTimeCallLogLabel = Label(updateCallLogFrame, text='Resolve time (hh:mm):')
updateResolveTimeCallLogLabel.grid(row=0, column=12, padx=10, pady=10)

updateResolveTimeCallLogEntry = Entry(updateCallLogFrame)
updateResolveTimeCallLogEntry.grid(row=0, column=13, padx=10, pady=10)

updateTimeTakenCallLogLabel = Label(updateCallLogFrame, text='Time taken (minutes):')
updateTimeTakenCallLogLabel.grid(row=0, column=15, padx=10, pady=10)

updateTimeTakenCallLogEntry = Entry(updateCallLogFrame)
updateTimeTakenCallLogEntry.grid(row=0, column=16, padx=10, pady=10)

updateResolveDescriptionCallLogLabel = Label(updateCallLogFrame, text='Resolve description:')
updateResolveDescriptionCallLogLabel.grid(row=1, column=0, padx=10, pady=10)

updateResolveDescriptionCallLogEntry = Entry(updateCallLogFrame)
updateResolveDescriptionCallLogEntry.grid(row=1, column=1, padx=10, pady=10)

#search for a specific issue id in call log function
def searchIssueIdCallLog():
    searchIssueIdCallLog = searchIssueIdCallLogEntry.get()
    
    for record in callLogTree.get_children():
        callLogTree.delete(record)
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT	    cl.issue_id AS 'Issue ID', 
                                e.employee_name AS 'Employee name', 
                                ho.operator_name AS 'Reporting operator', 
                                cl.call_time AS 'Call time', 
                                cl.call_date AS 'Call date',
                                a.asset_type AS 'Asset type', 
                                a.asset_make AS 'Asset make', 
                                a.operating_system AS 'OS', 
                                s.software_name AS 'Software', 
                                s.[valid_license?] AS 'Has a valid license?',
                                cl.issue_type AS 'Issue type', 
                                cl.issue_description AS 'Issue description', 
                                ho2.operator_name AS 'Assigned operator', 
                                cl.[issue_closed?] AS 'Closed?',
                                cl.closed_time AS 'Closed time', 
                                cl.closed_date AS 'Closed date', 
                                cl.resolution_description AS 'Resolution description', 
                                cl.minutes_taken_to_resolve AS 'Time taken to resolve (minutes)'
                    FROM        call_log cl
                    INNER JOIN  employee e 
                    ON          cl.employee_id = e.employee_id
                    INNER JOIN  helpdesk_operator ho 
                    ON          cl.operator_id = ho.operator_id
                    INNER JOIN  asset a 
                    ON          cl.asset_serial_number = a.serial_number
                    INNER JOIN  software s 
                    ON          cl.software_id = s.software_id
                    INNER JOIN  helpdesk_operator ho2 
                    ON          cl.assigned_operator = ho2.operator_id
                    WHERE       cl.issue_id LIKE ?''',
              ('%' + searchIssueIdCallLog + '%',))
    nameSearch = c.fetchall()
    data = 0
    
    for record in nameSearch:
        callLogTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5],
                                   record[6], record[7], record[8], record[9], record[10], record[11],
                                   record[12], record[13], record[14], record[15], record[16],
                                   record[17]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchEmployeeNameCallLogButton = Button(searchCallLogFrame,
                             text='Submit',
                             command=searchIssueIdCallLog)
searchEmployeeNameCallLogButton.grid(row=0, column=3, padx=10, pady=10)

#search for a specific employee in call log function
def searchEmployeeNameCallLog():
    searchEmployeeNameCallLog = searchEmployeeNameCallLogEntry.get()
    
    for record in callLogTree.get_children():
        callLogTree.delete(record)
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT	    cl.issue_id AS 'Issue ID', 
                                e.employee_name AS 'Employee name', 
                                ho.operator_name AS 'Reporting operator', 
                                cl.call_time AS 'Call time', 
                                cl.call_date AS 'Call date',
                                a.asset_type AS 'Asset type', 
                                a.asset_make AS 'Asset make', 
                                a.operating_system AS 'OS', 
                                s.software_name AS 'Software', 
                                s.[valid_license?] AS 'Has a valid license?',
                                cl.issue_type AS 'Issue type', 
                                cl.issue_description AS 'Issue description', 
                                ho2.operator_name AS 'Assigned operator', 
                                cl.[issue_closed?] AS 'Closed?',
                                cl.closed_time AS 'Closed time', 
                                cl.closed_date AS 'Closed date', 
                                cl.resolution_description AS 'Resolution description', 
                                cl.minutes_taken_to_resolve AS 'Time taken to resolve (minutes)'
                    FROM        call_log cl
                    INNER JOIN  employee e 
                    ON          cl.employee_id = e.employee_id
                    INNER JOIN  helpdesk_operator ho 
                    ON          cl.operator_id = ho.operator_id
                    INNER JOIN  asset a 
                    ON          cl.asset_serial_number = a.serial_number
                    INNER JOIN  software s 
                    ON          cl.software_id = s.software_id
                    INNER JOIN  helpdesk_operator ho2 
                    ON          cl.assigned_operator = ho2.operator_id
                    WHERE       e.employee_name LIKE ?''',
              ('%' + searchEmployeeNameCallLog + '%',))
    nameSearch = c.fetchall()
    data = 0
    
    for record in nameSearch:
        callLogTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5],
                                   record[6], record[7], record[8], record[9], record[10], record[11],
                                   record[12], record[13], record[14], record[15], record[16],
                                   record[17]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchEmployeeNameCallLogButton = Button(searchCallLogFrame,
                             text='Submit',
                             command=searchEmployeeNameCallLog)
searchEmployeeNameCallLogButton.grid(row=0, column=6, padx=10, pady=10)

#search for a specific issue type in call log function
def searchIssueTypeCallLog():
    searchIssueTypeCallLog = issueTypeVariable.get()
    
    for record in callLogTree.get_children():
        callLogTree.delete(record)
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT	    cl.issue_id AS 'Issue ID', 
                                e.employee_name AS 'Employee name', 
                                ho.operator_name AS 'Reporting operator', 
                                cl.call_time AS 'Call time', 
                                cl.call_date AS 'Call date',
                                a.asset_type AS 'Asset type', 
                                a.asset_make AS 'Asset make', 
                                a.operating_system AS 'OS', 
                                s.software_name AS 'Software', 
                                s.[valid_license?] AS 'Has a valid license?',
                                cl.issue_type AS 'Issue type', 
                                cl.issue_description AS 'Issue description', 
                                ho2.operator_name AS 'Assigned operator', 
                                cl.[issue_closed?] AS 'Closed?',
                                cl.closed_time AS 'Closed time', 
                                cl.closed_date AS 'Closed date', 
                                cl.resolution_description AS 'Resolution description', 
                                cl.minutes_taken_to_resolve AS 'Time taken to resolve (minutes)'
                    FROM        call_log cl
                    INNER JOIN  employee e 
                    ON          cl.employee_id = e.employee_id
                    INNER JOIN  helpdesk_operator ho 
                    ON          cl.operator_id = ho.operator_id
                    INNER JOIN  asset a 
                    ON          cl.asset_serial_number = a.serial_number
                    INNER JOIN  software s 
                    ON          cl.software_id = s.software_id
                    INNER JOIN  helpdesk_operator ho2 
                    ON          cl.assigned_operator = ho2.operator_id
                    WHERE       cl.issue_type LIKE ?''',
              ('%' + searchIssueTypeCallLog + '%',))
    nameSearch = c.fetchall()
    data = 0
    
    for record in nameSearch:
        callLogTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5],
                                   record[6], record[7], record[8], record[9], record[10], record[11],
                                   record[12], record[13], record[14], record[15], record[16],
                                   record[17]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchIssueTypeCallLogButton = Button(searchCallLogFrame,
                             text='Submit',
                             command=searchIssueTypeCallLog)
searchIssueTypeCallLogButton.grid(row=0, column=10, padx=10, pady=10)

#search for a specific asset make in call log function
def searchAssetMakeCallLog():
    searchAssetMakeCallLog = searchAssetMakeCallLogEntry.get()
    
    for record in callLogTree.get_children():
        callLogTree.delete(record)
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT	    cl.issue_id AS 'Issue ID', 
                                e.employee_name AS 'Employee name', 
                                ho.operator_name AS 'Reporting operator', 
                                cl.call_time AS 'Call time', 
                                cl.call_date AS 'Call date',
                                a.asset_type AS 'Asset type', 
                                a.asset_make AS 'Asset make', 
                                a.operating_system AS 'OS', 
                                s.software_name AS 'Software', 
                                s.[valid_license?] AS 'Has a valid license?',
                                cl.issue_type AS 'Issue type', 
                                cl.issue_description AS 'Issue description', 
                                ho2.operator_name AS 'Assigned operator', 
                                cl.[issue_closed?] AS 'Closed?',
                                cl.closed_time AS 'Closed time', 
                                cl.closed_date AS 'Closed date', 
                                cl.resolution_description AS 'Resolution description', 
                                cl.minutes_taken_to_resolve AS 'Time taken to resolve (minutes)'
                    FROM        call_log cl
                    INNER JOIN  employee e 
                    ON          cl.employee_id = e.employee_id
                    INNER JOIN  helpdesk_operator ho 
                    ON          cl.operator_id = ho.operator_id
                    INNER JOIN  asset a 
                    ON          cl.asset_serial_number = a.serial_number
                    INNER JOIN  software s 
                    ON          cl.software_id = s.software_id
                    INNER JOIN  helpdesk_operator ho2 
                    ON          cl.assigned_operator = ho2.operator_id
                    WHERE       a.asset_make LIKE ?''',
              ('%' + searchAssetMakeCallLog + '%',))
    nameSearch = c.fetchall()
    data = 0
    
    for record in nameSearch:
        callLogTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5],
                                   record[6], record[7], record[8], record[9], record[10], record[11],
                                   record[12], record[13], record[14], record[15], record[16],
                                   record[17]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchAssetMakeCallLogButton = Button(searchCallLogFrame,
                             text='Submit',
                             command=searchAssetMakeCallLog)
searchAssetMakeCallLogButton.grid(row=0, column=14, padx=10, pady=10)

#search for a specific operators open calls in call log function
def searchOperatorCallLog():
    searchOperatorCallLog = searchOperatorsCallLogEntry.get()
    
    for record in callLogTree.get_children():
        callLogTree.delete(record)
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT	    cl.issue_id AS 'Issue ID', 
                                e.employee_name AS 'Employee name', 
                                ho.operator_name AS 'Reporting operator', 
                                cl.call_time AS 'Call time', 
                                cl.call_date AS 'Call date',
                                a.asset_type AS 'Asset type', 
                                a.asset_make AS 'Asset make', 
                                a.operating_system AS 'OS', 
                                s.software_name AS 'Software', 
                                s.[valid_license?] AS 'Has a valid license?',
                                cl.issue_type AS 'Issue type', 
                                cl.issue_description AS 'Issue description', 
                                ho2.operator_name AS 'Assigned operator', 
                                cl.[issue_closed?] AS 'Closed?',
                                cl.closed_time AS 'Closed time', 
                                cl.closed_date AS 'Closed date', 
                                cl.resolution_description AS 'Resolution description', 
                                cl.minutes_taken_to_resolve AS 'Time taken to resolve (minutes)'
                    FROM        call_log cl
                    INNER JOIN  employee e 
                    ON          cl.employee_id = e.employee_id
                    INNER JOIN  helpdesk_operator ho 
                    ON          cl.operator_id = ho.operator_id
                    INNER JOIN  asset a 
                    ON          cl.asset_serial_number = a.serial_number
                    INNER JOIN  software s 
                    ON          cl.software_id = s.software_id
                    INNER JOIN  helpdesk_operator ho2 
                    ON          cl.assigned_operator = ho2.operator_id
                    WHERE       ho2.operator_name LIKE ?
                    AND         cl.[issue_closed?] = 0''',
                    ('%' + searchOperatorCallLog + '%',))
                    
    nameSearch = c.fetchall()
    data = 0
    
    for record in nameSearch:
        callLogTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5],
                                   record[6], record[7], record[8], record[9], record[10], record[11],
                                   record[12], record[13], record[14], record[15], record[16],
                                   record[17]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchOperatorCallLogButton = Button(searchOperatorCallLogFrame,
                             text='Submit',
                             command=searchOperatorCallLog)
searchOperatorCallLogButton.grid(row=0, column=2, padx=10, pady=10)

#update a specific call in call log function
def updateCallLog():
    issueToUpdateCallLog = issueToUpdateCallLogEntry.get().strip()
    updateIssueTypeCallLog = issueTypeVariable.get().strip()
    updateIssueClosedCallLog = updateIssueClosedCallLogEntry.get().strip()
    updateResolveDateCallLog = updateResolveDateCallLogEntry.get().strip()
    updateResolveTimeCallLog = updateResolveTimeCallLogEntry.get().strip()
    updateTimeTakenCallLog = updateTimeTakenCallLogEntry.get().strip()
    updateResolveDescriptionCallLog = updateResolveDescriptionCallLogEntry.get().strip()
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    if not updateIssueTypeCallLog or not updateIssueClosedCallLog or not issueToUpdateCallLog:
        messagebox.showerror('Mandatory fields', 'Fields marked * must be complete.')
        return
    try:
        c = conn.cursor()
        updateSql = ''' UPDATE      call_log
                        SET         issue_type = ?,
                                    [issue_closed?] = ?,
                                    closed_time = ?,
                                    closed_date = ?,
                                    resolution_description = ?,
                                    minutes_taken_to_resolve = ?
                        WHERE       issue_id = ?'''
        c.execute(updateSql, [updateIssueTypeCallLog, updateIssueClosedCallLog,
                            updateResolveTimeCallLog, updateResolveDateCallLog,
                            updateResolveDescriptionCallLog, updateTimeTakenCallLog,
                            issueToUpdateCallLog])
        conn.commit()
        messagebox.showinfo('Call log updated', 'Call log update successfully.')
    except Exception as e:
        messagebox.showerror('Error', f'An error has occured: {e}. Try again.')
    finally:
        conn.close()

    issueToUpdateCallLogEntry.delete(0, END)
    updateIssueClosedCallLogEntry.delete(0, END)
    updateResolveDateCallLogEntry.delete(0, END)
    updateResolveTimeCallLogEntry.delete(0, END)
    updateTimeTakenCallLogEntry.delete(0, END)
    updateResolveDescriptionCallLogEntry.delete(0, END)
    
updateCallLogButton = Button(updateCallLogFrame,
                             text='Submit',
                             command=updateCallLog)
updateCallLogButton.grid(row=1, column=3, padx=10, pady=10)

#
#Employee page content
#
employeePage = Frame(notebook)
notebook.add(employeePage, text='Employees')
l = Label(employeePage, text='Employee Information')
l.config(font=('Arial', 12))
l.pack(pady=10)

#configure treeview style
style = ttk.Style()
style.theme_use('default')
style.configure('Treeview',
                background='lightgray',
                foreground='black',
                rowheight=20,
                fieldbackground='lightgray')
style.map('Treeview', background=[('selected',
                                   'darkblue')])

#set up the treeview frame
employeeTreeFrame = Frame(employeePage)
employeeTreeFrame.pack(padx=5, pady=10, anchor='n')
employeeTreeScrollY = Scrollbar(employeeTreeFrame, orient='vertical')
employeeTreeScrollY.pack(side=RIGHT, fill=Y)
employeeTreeScrollX = Scrollbar(employeeTreeFrame, orient='horizontal')
employeeTreeScrollX.pack(side=BOTTOM, fill=X)

employeeTree = ttk.Treeview(employeeTreeFrame,
                           yscrollcommand=employeeTreeScrollY.set,
                           xscrollcommand=employeeTreeScrollX.set,
                           selectmode='extended')
employeeTree.pack()

#set up the treeview content
employeeTree['columns'] = ('employee_id',
                          'employee_name',
                          'job_title',
                          'department')
employeeTree.column('#0', width=0, stretch=NO)
employeeTree.column('employee_id', anchor=CENTER, width=100)
employeeTree.column('employee_name', anchor=CENTER, width=100)
employeeTree.column('job_title', anchor=CENTER, width=100)
employeeTree.column('department', anchor=CENTER, width=100)

employeeTree.heading('#0', text='', anchor=CENTER)
employeeTree.heading('employee_id', text='Employee ID' ,anchor=CENTER)
employeeTree.heading('employee_name', text='Employee Name', anchor=CENTER)
employeeTree.heading('job_title', text='Job Title', anchor=CENTER)
employeeTree.heading('department', text='Department', anchor=CENTER)

employeeTreeLabelFrame = LabelFrame(employeePage, text='View employee details here.')
employeeTreeLabelFrame.pack(fill='x', expand=False, padx=10, pady=2)

#employee search frame
searchEmployeeFrame = LabelFrame(employeePage,
                                text='Search for a specific employee here.')
searchEmployeeFrame.pack(fill='x', expand=False, padx=10, pady=2)

searchEmployeeLabel = Label(searchEmployeeFrame, text='Employee name:')
searchEmployeeLabel.grid(row=0, column=0, padx=10, pady=10)

searchEmployeeEntry = Entry(searchEmployeeFrame)
searchEmployeeEntry.grid(row=0, column=1, padx=10, pady=10)

#search for a specific employee function
def searchEmployee():
    searchEmployee = searchEmployeeEntry.get()
    
    for record in employeeTree.get_children():
        employeeTree.delete(record)
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT      e.employee_id AS 'Employee ID', 
		                        e.employee_name AS 'Employee Name', 
		                        jr.job_title AS 'Job Title', 
		                        jr.department AS 'Department'
                    FROM        employee e
                    INNER JOIN  job_role jr
                    ON          e.job_id = jr.job_id
                    WHERE       e.employee_name LIKE ?''',
              ('%' + searchEmployee + '%',))
    nameSearch = c.fetchall()
    data = 0
    
    for record in nameSearch:
        employeeTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2], record[3]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchEmployeeButton = Button(searchEmployeeFrame,
                             text='Submit',
                             command=searchEmployee)
searchEmployeeButton.grid(row=0, column=2, padx=10, pady=10)

#
#Operator page content
#
operatorPage = Frame(notebook)
notebook.add(operatorPage, text='Operators')
l = Label(operatorPage, text='Operator Information')
l.config(font=('Arial', 12))
l.pack(pady=10)

#configure treeview style
style = ttk.Style()
style.theme_use('default')
style.configure('Treeview',
                background='lightgray',
                foreground='black',
                rowheight=20,
                fieldbackground='lightgray')
style.map('Treeview', background=[('selected',
                                   'darkblue')])

#set up the treeview frame
operatorTreeFrame = Frame(operatorPage)
operatorTreeFrame.pack(padx=5, pady=10, anchor='n')
operatorTreeScrollY = Scrollbar(operatorTreeFrame, orient='vertical')
operatorTreeScrollY.pack(side=RIGHT, fill=Y)
operatorTreeScrollX = Scrollbar(operatorTreeFrame, orient='horizontal')
operatorTreeScrollX.pack(side=BOTTOM, fill=X)

operatorTree = ttk.Treeview(operatorTreeFrame,
                           yscrollcommand=operatorTreeScrollY.set,
                           xscrollcommand=operatorTreeScrollX.set,
                           selectmode='extended')
operatorTree.pack()

#set up the treeview content
operatorTree['columns'] = ('operator_id',
                          'operator_name',
                          'primary_specialisation',
                          'secondary_specialisation')
operatorTree.column('#0', width=0, stretch=NO)
operatorTree.column('operator_id', anchor=CENTER, width=100)
operatorTree.column('operator_name', anchor=CENTER, width=100)
operatorTree.column('primary_specialisation', anchor=CENTER, width=150)
operatorTree.column('secondary_specialisation', anchor=CENTER, width=150)

operatorTree.heading('#0', text='', anchor=CENTER)
operatorTree.heading('operator_id', text='Operator ID' ,anchor=CENTER)
operatorTree.heading('operator_name', text='Operator Name', anchor=CENTER)
operatorTree.heading('primary_specialisation', text='Primary Specialisation', anchor=CENTER)
operatorTree.heading('secondary_specialisation', text='Secondary Specialisation', anchor=CENTER)

operatorTreeLabelFrame = LabelFrame(operatorPage, text='View operator details here.')
operatorTreeLabelFrame.pack(fill='x', expand=False, padx=10, pady=2)

#search all operators frame
searchOperatorsFrame = LabelFrame(operatorPage,
                                text='Search for all operators here.')
searchOperatorsFrame.pack(fill='x', expand=False, padx=10, pady=2)

#search for all operators function
def searchOperators():
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT	    ho.operator_id AS 'Operator ID', 
                                ho.operator_name AS 'Operator Name', 
                                hs.specialisation_name AS 'Main Specialisation', 
                                hs2.specialisation_name AS 'Secondary Specialisaion'
                    FROM        dbo.helpdesk_operator ho
                    INNER JOIN  dbo.helpdesk_specialisation hs
                    ON          ho.primary_specialisation = hs.specialisation_id
                    INNER JOIN  dbo.helpdesk_specialisation hs2
                    ON          ho.secondary_specialisation = hs2. specialisation_id''')
    nameSearch = c.fetchall()
    data = 0
    
    for record in nameSearch:
        operatorTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2], record[3]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchOperatorButton = Button(searchOperatorsFrame,
                             text='Search all',
                             command=searchOperators)
searchOperatorButton.grid(row=0, column=0, padx=10, pady=10)

#
#Asset page content
#
assetPage = Frame(notebook)
notebook.add(assetPage, text='IT Assets')
l = Label(assetPage, text='IT Asset Information')
l.config(font=('Arial', 12))
l.pack(pady=10)

#configure treeview style
style = ttk.Style()
style.theme_use('default')
style.configure('Treeview',
                background='lightgray',
                foreground='black',
                rowheight=20,
                rowwidth=30,
                fieldbackground='lightgray')
style.map('Treeview', background=[('selected',
                                   'darkblue')])

#set up the treeview frame
assetTreeFrame = Frame(assetPage)
assetTreeFrame.pack(padx=5, pady=10, anchor='n')
assetTreeScrollY = Scrollbar(assetTreeFrame, orient='vertical')
assetTreeScrollY.pack(side=RIGHT, fill=Y)
assetTreeScrollX = Scrollbar(assetTreeFrame, orient='horizontal')
assetTreeScrollX.pack(side=BOTTOM, fill=X)

assetTree = ttk.Treeview(assetTreeFrame,
                           yscrollcommand=assetTreeScrollY.set,
                           xscrollcommand=assetTreeScrollX.set,
                           selectmode='extended')
assetTree.pack()

#set up the treeview content
assetTree['columns'] = ('serial_number',
                        'asset_type',
                        'asset_make',
                        'operating_system')
assetTree.column('#0', width=0, stretch=NO)
assetTree.column('serial_number', anchor=CENTER, width=100)
assetTree.column('asset_type', anchor=CENTER, width=100)
assetTree.column('asset_make', anchor=CENTER, width=100)
assetTree.column('operating_system', anchor=CENTER, width=100)

assetTree.heading('#0', text='', anchor=CENTER)
assetTree.heading('serial_number', text='Serial Number' ,anchor=CENTER)
assetTree.heading('asset_type', text='Asset Type', anchor=CENTER)
assetTree.heading('asset_make', text='Asset Make', anchor=CENTER)
assetTree.heading('operating_system', text='Operating System', anchor=CENTER)

assetTreeLabelFrame = LabelFrame(assetPage, text='View IT asset details here.')
assetTreeLabelFrame.pack(fill='x', expand=False, padx=10, pady=2)

#search for a specific asset frame
searchAssetFrame = LabelFrame(assetPage,
                                text='Search for specific assets here.')
searchAssetFrame.pack(fill='x', expand=False, padx=10, pady=2)

searchAssetLabel = Label(searchAssetFrame, text='Asset serial number:')
searchAssetLabel.grid(row=0, column=0, padx=10, pady=10)

searchAssetEntry = Entry(searchAssetFrame)
searchAssetEntry.grid(row=0, column=1, padx=10, pady=10)

#Search for a specific asset function
def searchAsset():
    searchAsset = searchAssetEntry.get()
    
    for record in assetTree.get_children():
        assetTree.delete(record)
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT	a.serial_number AS 'Serial number',
		                    a.asset_type AS 'Asset type',
		                    a.asset_make AS 'Asset make',
		                    a.operating_system AS 'OS'
                    FROM    dbo.asset a
                    WHERE   a.serial_number LIKE ?''',
              ('%' + searchAsset + '%',))
    assetSearch = c.fetchall()
    data = 0
    
    for record in assetSearch:
        assetTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2], record[3]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchAssetButton = Button(searchAssetFrame,
                             text='Submit',
                             command=searchAsset)
searchAssetButton.grid(row=0, column=2, padx=10, pady=10)

#
#Software page content
#
softwarePage = Frame(notebook)
notebook.add(softwarePage, text='IT Software')
l = Label(softwarePage, text='IT Software Information')
l.config(font=('Arial', 12))
l.pack(pady=10)

#configure treeview style
style = ttk.Style()
style.theme_use('default')
style.configure('Treeview',
                background='lightgray',
                foreground='black',
                rowheight=20,
                rowwidth=30,
                fieldbackground='lightgray')
style.map('Treeview', background=[('selected',
                                   'darkblue')])

#set up the treeview frame
softwareTreeFrame = Frame(softwarePage)
softwareTreeFrame.pack(padx=5, pady=10, anchor='n')
softwareTreeScrollY = Scrollbar(softwareTreeFrame, orient='vertical')
softwareTreeScrollY.pack(side=RIGHT, fill=Y)
softwareTreeScrollX = Scrollbar(softwareTreeFrame, orient='horizontal')
softwareTreeScrollX.pack(side=BOTTOM, fill=X)

softwareTree = ttk.Treeview(softwareTreeFrame,
                           yscrollcommand=softwareTreeScrollY.set,
                           xscrollcommand=softwareTreeScrollX.set,
                           selectmode='extended')
softwareTree.pack()

#set up the treeview content
softwareTree['columns'] = ('software_id',
                        'software_name',
                        'valid_license?')
softwareTree.column('#0', width=0, stretch=NO)
softwareTree.column('software_id', anchor=CENTER, width=100)
softwareTree.column('software_name', anchor=CENTER, width=100)
softwareTree.column('valid_license?', anchor=CENTER, width=100)

softwareTree.heading('#0', text='', anchor=CENTER)
softwareTree.heading('software_id', text='Software ID' ,anchor=CENTER)
softwareTree.heading('software_name', text='Software Name', anchor=CENTER)
softwareTree.heading('valid_license?', text='Valid License?', anchor=CENTER)

softwareTreeLabelFrame = LabelFrame(softwarePage, text='View IT software details here.')
softwareTreeLabelFrame.pack(fill='x', expand=False, padx=10, pady=2)

#search for a specific software frame
searchSoftwareFrame = LabelFrame(softwarePage,
                                text='Search for specific software here.')
searchSoftwareFrame.pack(fill='x', expand=False, padx=10, pady=2)

searchSoftwareLabel = Label(searchSoftwareFrame, text='Software name:')
searchSoftwareLabel.grid(row=0, column=0, padx=10, pady=10)

searchSoftwareEntry = Entry(searchSoftwareFrame)
searchSoftwareEntry.grid(row=0, column=1, padx=10, pady=10)

#Search for a specific software function
def searchSoftware():
    searchSoftware = searchSoftwareEntry.get()
    
    for record in softwareTree.get_children():
        softwareTree.delete(record)
    
    conn = db.connect('Driver={SQL Server};'
                  'Server=LAPTOP-JNA8CL44\\SQLEXPRESS;'
                  'Database=manzaneque_ltd;'
                  'Trusted_Connection=yes;')
    
    c = conn.cursor()
    c.execute('''   SELECT	s.software_id AS 'Software ID',
		                    s.software_name AS 'Software name',
		                    s.[valid_license?] AS 'Has a valid license?'
                    FROM    dbo.software s
                    WHERE   s.software_name LIKE ?''',
              ('%' + searchSoftware + '%',))
    softwareSearch = c.fetchall()
    data = 0
    
    for record in softwareSearch:
        softwareTree.insert(parent='',
                           index='end',
                           iid=data,
                           text='',
                           values=(record[0], record[1], record[2]))
        data += 1
    
    conn.commit()
    conn.close()
    
searchSoftwareButton = Button(searchSoftwareFrame,
                             text='Submit',
                             command=searchSoftware)
searchSoftwareButton.grid(row=0, column=2, padx=10, pady=10)

#finish initialisation of tkinter gui elements
root.mainloop()