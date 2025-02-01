import pyodbc as db
import pandas as pd
from tkinter import ttk
from tkinter import *

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