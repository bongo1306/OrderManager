import wx
#import wx.xrc
from wx import xrc
ctrl = xrc.XRCCTRL
import wx.calendar
import wx.lib.inspection
from win32com.client import Dispatch
from dateutil.parser import parse
#import numpy as np
#import easygui
import pyodbc
import re
import datetime
import General as gn
import Database as db
import threading
import sys
import os
import csv
import os
import time
import re
import datetime
import string
import xlrd

class AdminTab(object):

        def init_admin_tab(self):                               
                              
                # Connect Events
                 
                self.name=wx.FindWindowByName('choice:name')
                self.notebook =wx.FindWindowByName('notebook:applications')
                #self.m_panel9=wx.FindWindowByName('m_panel9')
                #self.m_panel9 = wx.Panel( self.applications)
                self.createuserpanel=wx.FindWindowByName('createuserpanel')
                self.m_staticName = wx.FindWindowByName('m_staticName')
                self.m_textName = wx.FindWindowByName('m_textName')   
                self.m_staticdepartment = wx.FindWindowByName('m_staticdepartment')
                self.m_combodepartment = wx.FindWindowByName('m_combodepartment')
                self.m_staticEmail = wx.FindWindowByName('m_staticEmail')
                self.m_textEmail = wx.FindWindowByName('m_textEmail')
                self.m_staticPassword = wx.FindWindowByName('m_staticPassword')
                self.m_textpassword = wx.FindWindowByName('m_textpassword')
                self.m_staticPasswordexpiration = wx.FindWindowByName('m_staticPasswordexpiration')
                self.m_datePickerpasswordexp = wx.FindWindowByName('m_datePickerpasswordexp')
                self.m_checkBoxactivated = wx.FindWindowByName('m_checkBoxactivated')
                self.m_checkBoxgetrevision = wx.FindWindowByName('m_checkBoxgetrevision')
                self.m_checkBoxgetpopsheet = wx.FindWindowByName('m_checkBoxgetpopsheet')
                self.m_checkBoxengineer = wx.FindWindowByName('m_checkBoxengineer')
                self.m_checkBoxcadeng = wx.FindWindowByName('m_checkBoxcadeng')
                self.m_checkBoxprojecteng = wx.FindWindowByName('m_checkBoxprojecteng')
                self.m_checkBoxprojectlead = wx.FindWindowByName('m_checkBoxprojectlead')
                self.m_checkBoxscheduler = wx.FindWindowByName('m_checkBoxscheduler')
                self.m_checkBoxapprovefirst = wx.FindWindowByName('m_checkBoxapprovefirst')
                self.m_checkBoxApprovesecond = wx.FindWindowByName('m_checkBoxApprovesecond')
                self.m_checkBoxReadonly = wx.FindWindowByName('m_checkBoxReadonly')
                self.m_checkBoxAdmin = wx.FindWindowByName('m_checkBoxAdmin')
                self.m_buttonSubmit = wx.FindWindowByName('m_buttonSubmit')
                self.m_export_to_excel=wx.FindWindowByName('m_export_to_excel')
                self.m_ButtonHolidays = wx.FindWindowByName('m_ButtonHolidays')
                self.m_DateAddHolidays = wx.FindWindowByName('m_DateAddHolidays')
                self.m_ButtonDisplayList = wx.FindWindowByName('m_ButtonDisplayList')
                self.m_TextCtrlHoliList = wx.FindWindowByName('m_TextCtrlHoliList')
                self.m_ButtonRemoveHoliday = wx.FindWindowByName('m_ButtonRemoveHoliday')
                self.m_DateRemoveHoliday = wx.FindWindowByName('m_DateRemoveHoliday')

                # Bind Do_Nothing Event upon mousewheel scroll in order to not change users Dropdowns selection accidently
                ctrl(self, 'm_combodepartment').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing3)

                self.fetchDB()
                adminuser=[]
                selected_user=self.name.GetStringSelection()
                self.dbCursor.execute('select name from employees where Admin=1')
                DB_values=self.dbCursor.fetchone()
                while DB_values!= None:
                        adminuser.append(DB_values[0])
                        DB_values=self.dbCursor.fetchone()
                
                if selected_user not in adminuser:                         
                         self.createuserpanel.Show(False)                         
                         self.m_staticName.Hide()
                         self.m_textName.Hide()
                         self.m_staticdepartment.Hide()
                         self.m_combodepartment.Hide()
                         self.m_staticEmail.Hide()
                         self.m_textEmail.Hide()
                         self.m_staticPassword.Hide()
                         self.m_textpassword.Hide()
                         self.m_staticPasswordexpiration.Hide()
                         self.m_datePickerpasswordexp.Hide()
                         self.m_checkBoxactivated.Hide()
                         self.m_checkBoxgetrevision.Hide()
                         self.m_checkBoxgetpopsheet.Hide()
                         self.m_checkBoxengineer.Hide()
                         self.m_checkBoxcadeng.Hide()
                         self.m_checkBoxprojecteng.Hide()
                         self.m_checkBoxprojectlead.Hide()
                         self.m_checkBoxscheduler.Hide()
                         self.m_checkBoxapprovefirst.Hide()
                         self.m_checkBoxApprovesecond.Hide()
                         self.m_checkBoxReadonly.Hide()
                         self.m_checkBoxAdmin.Hide()
                         self.m_buttonSubmit.Hide()
                         self.m_export_to_excel.Hide()
                         self.m_ButtonHolidays.Hide()
                         self.m_DateAddHolidays.Hide()
                         self.m_ButtonDisplayList.Hide()
                         self.m_TextCtrlHoliList.Hide()
                         self.m_ButtonRemoveHoliday.Hide()
                         self.m_DateRemoveHoliday.Hide()

                else:
                        
                        self.createuserpanel=wx.FindWindowByName('createuserpanel')
                        self.Bind(wx.EVT_BUTTON, self.insertdetails, id=xrc.XRCID('m_buttonSubmit'))
                        self.Bind(wx.EVT_BUTTON, self.OnBtnExportData, id=xrc.XRCID('m_export_to_excel'))
                        self.Bind(wx.EVT_BUTTON, self.OnBtnAddHolidays, id=xrc.XRCID('m_ButtonHolidays'))
                        self.Bind(wx.EVT_BUTTON, self.DisplayHolidayList, id=xrc.XRCID('m_ButtonDisplayList'))
                        self.Bind(wx.EVT_BUTTON, self.OnBtnRemoveHoliday, id=xrc.XRCID('m_ButtonRemoveHoliday'))
                        self.m_textName = wx.FindWindowByName('m_textName')                
                        self.m_combodepartment = wx.FindWindowByName('m_combodepartment')
                        self.m_textEmail = wx.FindWindowByName('m_textEmail')
                        self.m_textpassword= wx.FindWindowByName('m_textpassword')
                        self.m_datePickerpasswordexp=wx.FindWindowByName('m_datePickerpasswordexp')
                        self.m_checkBoxactivated=wx.FindWindowByName('m_checkBoxactivated')
                        self.m_checkBoxgetrevision=wx.FindWindowByName('m_checkBoxgetrevision')
                        self.m_checkBoxgetpopsheet=wx.FindWindowByName('m_checkBoxgetpopsheet')
                        self.m_checkBoxengineer=wx.FindWindowByName('m_checkBoxengineer')
                        self.m_checkBoxcadeng=wx.FindWindowByName('m_checkBoxcadeng')
                        self.m_checkBoxprojecteng=wx.FindWindowByName('m_checkBoxprojecteng')
                        self.m_checkBoxprojectlead=wx.FindWindowByName('m_checkBoxprojectlead')
                        self.m_checkBoxscheduler=wx.FindWindowByName('m_checkBoxscheduler')
                        self.m_checkBoxapprovefirst=wx.FindWindowByName('m_checkBoxapprovefirst')
                        self.m_checkBoxApprovesecond=wx.FindWindowByName('m_checkBoxApprovesecond')
                        self.m_checkBoxReadonly=wx.FindWindowByName('m_checkBoxReadonly')
                        self.m_checkBoxAdmin = wx.FindWindowByName('m_checkBoxAdmin')
                        self.m_buttonSubmit=wx.FindWindowByName('m_buttonSubmit')
                        self.m_export_to_excel=wx.FindWindowByName('m_export_to_excel')
                        self.m_ButtonHolidays = wx.FindWindowByName('m_ButtonHolidays')
                        self.m_DateAddHolidays = wx.FindWindowByName('m_DateAddHolidays')
                        self.m_ButtonDisplayList = wx.FindWindowByName('m_ButtonDisplayList')
                        self.m_TextCtrlHoliList = wx.FindWindowByName('m_TextCtrlHoliList')
                        self.m_ButtonRemoveHoliday = wx.FindWindowByName('m_ButtonRemoveHoliday')
                        self.m_DateRemoveHoliday = wx.FindWindowByName('m_DateRemoveHoliday')
                        
                        #Database connection variables
                        self.AENames = []
                        #self.dbCursor = None

        def do_nothing3(self, evt):
            print 'on events pit'
                        
        def fetchDB(self):

                #Database connection variables
                self.conn = db.connect_to_eng04_database() #create a new database connection for admin.                
                self.dbCursor = self.conn.cursor()
                self.db_dropdown()
               
                
        

        #method for binding dropdown with departments
        def db_dropdown(self):
                
              del self.AENames[:]                           
              self.dbCursor.execute('select department from departments')
              row = self.dbCursor.fetchone()
              
              while row != None:                
                      self.AENames.append(row[0])
                      row = self.dbCursor.fetchone()
                      
              self.m_combodepartment.SetItems(self.AENames)             


        #method for insert or update record to db
        def insertdetails(self,event):        
           activated=self.m_checkBoxactivated.GetValue()
           name=str(self.m_textName.GetValue())
           password= str(self.m_textpassword.GetValue())           
           passwordexp=self.m_datePickerpasswordexp.GetValue()                      
           dt=parse(str(passwordexp))           
           email=str(self.m_textEmail.GetValue())           
           department=str(self.m_combodepartment.GetValue())           
           revision_notice=self.m_checkBoxgetrevision.GetValue()           
           popsheet=self.m_checkBoxgetpopsheet.GetValue()           
           isenginner=self.m_checkBoxengineer.GetValue()           
           cad=self.m_checkBoxcadeng.GetValue()           
           projectenginner=self.m_checkBoxprojecteng.GetValue()
           projectlead=self.m_checkBoxprojectlead.GetValue()
           scheduler=self.m_checkBoxscheduler.GetValue()         
           approvefirst=self.m_checkBoxapprovefirst.GetValue()           
           approvesecond=self.m_checkBoxApprovesecond.GetValue()          
           readonly= self.m_checkBoxReadonly.GetValue()
           admin=self.m_checkBoxAdmin.GetValue()
           self.dbCursor.execute("select max(id) from  [employees]")           
           ID=self.dbCursor.fetchone()                      
           u_ID=int(ID[0])                    
           id=u_ID+1           
           self.dbCursor.execute("select * from  [employees] where email = ?",(email,))                    
           listvalue=self.dbCursor.fetchone()           
           if listvalue:
                         if not name or not department or not email or not password:
                               msgbox = wx.MessageBox('Please enter your details', 'Alert')
                         else:
                                 if passwordexp.IsValid() == True and passwordexp.IsLaterThan(wx.DateTime_Today()) == False:      
                                           msgbox = wx.MessageBox('The password expiration date cannot be in past!', 'Alert')
                                 else:
                                         
                                       if ('@LENNOXINTL.COM' in email.upper()) or ('@HEATCRAFTRPD.COM' in email.upper()) or ('@KYSORWARREN.COM' in email.upper()):
                                               regex=re.match('^[_a-zA-Z0-9-]+(\.[_a-zA-Z0-9-]+)*@[a-zA-Z0-9-]+(\.[a-zA-Z]{2,3})$',email.upper())
                                               if regex:
                                                        self.dbCursor.execute("""UPDATE employees SET activated=?,name=?,password=?,password_expiration=?,
                                                       email=?,department=?,
                                                       gets_revision_notice=?
                                                      ,gets_pop_sheet_notice=?
                                                      ,is_engineer=?
                                                      ,is_cad_designer=?
                                                      ,is_project_engineer=?
                                                      ,is_project_lead=?
                                                      ,is_scheduler=?
                                                      ,can_approve_first=?
                                                      ,can_approve_second=?
                                                      ,is_readonly=?,
                                                      Admin=?
                                                      WHERE email=?""",
                                                      (activated,name, password,dt,email,department,revision_notice,popsheet,isenginner,cad,projectenginner,projectlead,scheduler,approvefirst,approvesecond,readonly,admin,email))
                                                        self.conn.commit() 
                                                        msgbox = wx.MessageBox('Data successfully updated', 'Alert')
                                               else:
                                                        msgbox = wx.MessageBox('Email should be like example@Lennoxintl.com or example@Heatcraftrpd.com or example@Kysorwarren.com', 'Alert')
                                                        
                                       else:
                                               msgbox = wx.MessageBox('Email should be like example@Lennoxintl.com or example@Heatcraftrpd.com or example@Kysorwarren.com', 'Alert')


           else:
                   if  passwordexp.IsValid() == True and passwordexp.IsLaterThan(wx.DateTime_Today()) == False:      
                                           msgbox = wx.MessageBox('The password expiration date cannot be in past!', 'Alert')
                                    
                   else:
                           if ('@LENNOXINTL.COM' in email.upper()) or ('HEATCRAFTRPD.COM' in email.upper()) or ('KYSORWARREN.COM' in email.upper()):
                                   regex=re.match('^[_a-zA-Z0-9-]+(\.[_a-zA-Z0-9-]+)*@[a-zA-Z0-9-]+(\.[a-zA-Z]{2,3})$',email.upper())
                                   if regex:
                                           
                                               self.dbCursor.execute("""insert into employees(activated,name,password,password_expiration,email,department,gets_revision_notice
                                              ,[gets_pop_sheet_notice]
                                              ,[is_engineer]
                                              ,[is_cad_designer]
                                              ,[is_project_engineer]
                                              ,[is_project_lead]
                                              ,[is_scheduler]
                                              ,[can_approve_first]
                                              ,[can_approve_second]
                                              ,[is_readonly],[Admin]) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                                             [activated,name, password,dt,email,department,revision_notice,popsheet,isenginner,cad,projectenginner,projectlead,scheduler,approvefirst,approvesecond,readonly,admin])                                                  
                                               self.conn.commit()
                                               msgbox = wx.MessageBox('Data successfully inserted', 'Alert')
                                   else:
                                        msgbox = wx.MessageBox('Email should be like example@Lennoxintl.com or example@Heatcraftrpd.com or example@Kysorwarren.com', 'Alert')
                           else:
                                   msgbox = wx.MessageBox('Email should be like example@Lennoxintl.com or example@Heatcraftrpd.com or example@Kysorwarren.com', 'Alert')

                        
        #method for export the excel of employee details                     
        def OnBtnExportData(self,event=None):
                excel = Dispatch('Excel.Application')
                excel.Visible = True
                wb = excel.Workbooks.Add()                
                self.dbCursor.execute("""
                SELECT Admin,activated AS Activated,
                name AS [Name],
                password AS[Password],
                password_expiration AS [Password Expiration],
                email AS [Email],
                department AS [Department],gets_revision_notice AS [Gets Revision Notice],gets_pop_sheet_notice as [Get Pop Sheet Notice],is_engineer as [Engineer],
                is_cad_designer as [CAD],is_project_engineer as [Project Engineer],is_project_lead as [Project Lead],is_scheduler as [Scheduler],can_approve_first as [Approve First],
                can_approve_second as [Approve Second],is_readonly as [Readonly] FROM  [employees] """)
                column=[ i[0] for i in self.dbCursor .description ]
                rows=self.dbCursor .fetchall()
                print len(column)
                print len(rows)
                    
                wb.ActiveSheet.Cells(1, 1).Value = rows
                wb.ActiveSheet.Columns(1).AutoFit()
                R=len(rows)
                C=len(column)
                 #Write the header
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(1,C))
                excel_range.Value = column
                
                #specify the excel range that the main data covers
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(2, 1), wb.ActiveSheet.Cells(R+1,C))
                excel_range.Value = rows

                #Autofit the columns
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(R+1,C))
                excel_range.Columns.AutoFit()
                
               
##                with open(save_path, 'wb') as csvfile:                                                                
##                        writer = csv.writer(csvfile)                        
##                        writer.writerow([ i[0] for i in self.dbCursor.description ]) # heading row                        
##                        writer.writerows(self.dbCursor.fetchall())
##                        
##                msgbox = wx.MessageBox('Export successfully', 'Alert')                
        def OnBtnAddHolidays(self, event = None):
            if self.m_DateAddHolidays.GetValue():
                holidaydate = str(self.m_DateAddHolidays.GetValue())
                holidaydate1 = datetime.datetime.strptime(holidaydate, "%d/%m/%Y %H:%M:%S").strftime("%m/%d/%Y %H:%M:%S")
                #print holidaydate1

                #current holiday list for checking
                currHolidaylist = 'SELECT * from Holidays'
                holidays = self.dbCursor.execute(currHolidaylist).fetchall()
                holidays_list = []
                for holiday in holidays:
                    for date in holiday:
                        holidays_list.append(date)

                #print holidays_list

                if holidaydate1 not in holidays_list:
                   self.dbCursor.execute("insert into Holidays (Holidays) values (?)", str(holidaydate1))
                   self.conn.commit()
                   msgbox = wx.MessageBox('Holiday successfully added', 'Alert')
                else:
                    msgbox = wx.MessageBox('Holiday already exists in database!', 'Alert')




        def DisplayHolidayList(self, event = None):
            currHolidaylist = 'SELECT * from Holidays ORDER by Holidays ASC'
            holidays = self.dbCursor.execute(currHolidaylist).fetchall()
            holidays_list = []
            for holiday in holidays:
                for date in holiday:
                    dateclean = str(date)

                    holidays_list.append(dateclean)

            datesdisplay = ''
            for dateclean in holidays_list:
                dateclean1 = string.replace(dateclean,'00:00:00','')
                datesdisplay += dateclean1
                datesdisplay += '\n'
                datesdisplay += '\n'

            self.m_TextCtrlHoliList.SetValue(str(datesdisplay))

            #SO = [[51462595.0], [51479183.0], [51479772.0], [51479772.0], [51479772.0], [51479772.0], [51479772.0], [51480692.0], [51490982.0], [51511816.0], [51511816.0], [51511816.0], [51511816.0], [51511816.0], [51511816.0], [51517023.0], [51518919.0], [51518919.0], [51522332.0], [51522332.0], [51522332.0], [51522332.0], [51522332.0], [51532548.0], [51532548.0], [51544399.0], [51544399.0], [51548610.0], [51549288.0], [51549288.0], [51549957.0], [51552838.0], [51554835.0], [51557330.0], [51567609.0], [51570493.0], [51570610.0], [51570610.0], [51570610.0], [51570796.0], [51570906.0], [51571087.0], [51571980.0], [51571980.0], [51571980.0], [51572944.0], [51573040.0], [51573913.0], [51573913.0], [51574920.0], [51574920.0], [51574920.0], [51575311.0], [51577249.0], [51577963.0], [51577963.0], [51578642.0], [51588535.0], [51589824.0], [51590214.0], [51590214.0], [51591619.0], [51593129.0], [51593949.0], [51593949.0], [51593985.0], [51593985.0], [51593987.0], [51594839.0], [51596139.0], [51596245.0], [51596246.0], [51596707.0], [51596852.0], [51596852.0], [51597013.0], [51597013.0], [51598792.0], [51598792.0], [51598851.0], [51598851.0], [51600307.0], [51600307.0], [51600307.0], [51601268.0], [51601268.0], [51601726.0], [51602258.0], [51602258.0], [51602347.0], [51602610.0], [51602801.0], [51602801.0], [51602809.0], [51603053.0], [51603311.0], [51603311.0], [51603311.0], [51603311.0], [51603855.0], [51603913.0], [51603913.0], [51604436.0], [51604542.0], [51604542.0], [51605167.0], [51605167.0], [51605462.0], [51605463.0], [51605493.0], [51605493.0], [51605679.0], [51605679.0], [51606219.0], [51606386.0], [51606635.0], [51606923.0], [51606923.0], [51607184.0], [51607185.0], [51607234.0], [51607267.0], [51607578.0], [51607578.0], [51607578.0], [51607578.0], [51607583.0], [51607814.0], [51607814.0], [51608237.0], [51608529.0], [51608564.0], [51608564.0], [51608564.0], [51608824.0], [51608824.0], [51608883.0], [51608883.0], [51609080.0], [51609080.0], [51609236.0], [51609270.0], [51610091.0], [51610092.0], [51610094.0], [51610094.0], [51610316.0], [51610810.0], [51610810.0], [51611061.0], [51611061.0], [51611061.0], [51611061.0], [51611061.0], [51612709.0], [51612914.0], [51612948.0], [51612948.0], [51613835.0], [51614060.0], [51615749.0], [51617240.0], [51617990.0], [51617990.0], [51617990.0], [51617990.0], [51617990.0], [51617990.0], [51617990.0], [51617990.0], [51617990.0], [51618378.0], [51618495.0], [51618495.0], [51619096.0], [51619096.0], [51619096.0], [51619835.0], [51619835.0], [51619835.0], [51619835.0], [51619835.0], [51619835.0], [51619835.0], [51619835.0], [51625542.0], [51625542.0], [51626994.0], [51626994.0], [51628039.0], [51628039.0], [51628039.0], [51628039.0], [51629047.0], [51629129.0], [51629129.0], [51629129.0], [51629129.0], [51629129.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51629382.0], [51631748.0], [51631748.0], [51631748.0], [51631748.0], [51631748.0], [51631748.0], [51635239.0], [51635239.0], [51635455.0], [51635455.0], [51637369.0], [51639045.0], [51639703.0], [51639703.0], [51640905.0], [51640905.0], [51644630.0], [51644630.0], [51649782.0], [51649782.0], [51651113.0], [51651957.0], [51651957.0], [51651957.0], [51651957.0], [51651957.0], [51651957.0], [51651957.0], [51651957.0], [51652131.0], [51653897.0], [51653897.0], [51654110.0], [51654110.0], [51654110.0], [51656689.0], [51656689.0], [51658171.0], [51659377.0], [51659377.0], [51659377.0], [51659555.0], [51659555.0], [51659555.0], [51660939.0], [51661601.0], [51662234.0], [51663271.0], [51663271.0], [51663271.0], [51663511.0], [51663628.0], [51663628.0], [51663628.0], [51663628.0], [51663628.0], [51663628.0], [51664067.0], [51664067.0], [51664067.0], [51664067.0], [51664067.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664943.0], [51664966.0], [51665123.0], [51665123.0], [51667941.0], [51667941.0], [51668520.0], [51669107.0], [51669107.0], [51669146.0], [51670106.0], [51670106.0], [51670106.0], [51670106.0], [51670106.0], [51670106.0], [51670106.0], [51670821.0], [51671707.0], [51672178.0], [51673692.0], [51673692.0], [51673692.0], [51673764.0], [51673785.0], [51674099.0], [51674099.0], [51674925.0], [51675132.0], [51675132.0], [51675346.0], [51675346.0], [51675346.0], [51675346.0], [51676142.0], [51676284.0], [51676284.0], [51676463.0], [51676617.0], [51676617.0], [51677845.0], [51677876.0], [51677876.0], [51678296.0], [51678296.0], [51679658.0], [51679658.0], [51679658.0], [51680116.0], [51680530.0], [51681035.0], [51681035.0], [51681611.0], [51681611.0], [51682200.0], [51682666.0], [51682666.0], [51682703.0], [51684188.0], [51684188.0], [51684220.0], [51684761.0], [51684761.0], [51684804.0], [51684859.0], [51685000.0], [51685146.0], [51687006.0], [51687006.0], [51687006.0], [51687006.0], [51687006.0], [51687006.0], [51687006.0], [51687006.0], [51687006.0], [51687006.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51687508.0], [51692517.0], [51692517.0], [51692517.0], [51692517.0], [51692517.0], [51692517.0], [51692517.0], [51693335.0], [51693335.0], [51693578.0], [51693578.0], [51693578.0], [51696045.0], [51696045.0], [51696045.0], [51696045.0], [51696045.0], [51696045.0], [51696438.0], [51696438.0], [51696716.0], [51699920.0], [51701064.0], [51701517.0], [51702600.0], [51702600.0], [51702600.0], [51702600.0], [51702600.0], [51702600.0], [51702600.0], [51702600.0], [51702600.0], [51702600.0], [51703600.0], [51705781.0], [51705781.0], [51705781.0], [51705781.0], [51705781.0], [51705848.0], [51705848.0], [51705848.0], [51705848.0], [51705848.0], [51705848.0], [51705848.0], [51705848.0], [51705848.0], [51706656.0], [51706656.0], [51706656.0], [51706656.0], [51706656.0], [51707406.0], [51707406.0], [51707406.0], [51707406.0], [51707406.0], [51707406.0], [51707406.0], [51708270.0], [51709901.0], [51709901.0], [51709901.0], [51709901.0], [51709901.0], [51709901.0], [51709901.0], [51709901.0], [51709901.0], [51717012.0], [51718990.0], [51718990.0], [51720665.0], [51720665.0], [51720665.0], [51721863.0], [51721863.0], [51721966.0], [51721966.0], [51721966.0], [51721966.0], [51721966.0], [51721966.0], [51723065.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51727917.0], [51728382.0], [51728382.0], [51728382.0], [51730628.0], [51730628.0], [51731523.0], [51731523.0], [51731837.0], [51731837.0], [51733194.0], [51733865.0], [51733865.0], [51734288.0], [51734288.0], [51734288.0], [51734342.0], [51734342.0], [51736518.0], [51737157.0], [51737852.0], [51737852.0], [51738135.0], [51738135.0], [51738599.0], [51738599.0], [51739172.0], [51739186.0], [51739186.0], [51739376.0], [51739376.0], [51739674.0], [51739674.0], [51741847.0], [51741847.0], [51741964.0], [51742349.0], [51742349.0], [51742485.0], [51742485.0], [51742649.0], [51742649.0], [51742666.0], [51742666.0], [51743026.0], [51743026.0], [51743077.0], [51744284.0], [51744284.0], [51744475.0], [51744475.0], [51745300.0], [51745300.0], [51745912.0], [51745912.0], [51745961.0], [51746182.0], [51746182.0], [51746586.0], [51746749.0], [51746749.0], [51747470.0], [51747539.0], [51747539.0], [51747758.0], [51747761.0], [51747904.0], [51747904.0], [51747904.0], [51747904.0], [51747904.0], [51747904.0], [51747904.0], [51747904.0], [51747904.0], [51747904.0], [51747904.0], [51748106.0], [51748746.0], [51748746.0], [51749082.0], [51749082.0], [51750123.0], [51750721.0], [51750721.0], [51750737.0], [51750737.0], [51750946.0], [51750991.0], [51750991.0], [51751041.0], [51751101.0], [51751269.0], [51751269.0], [51753747.0], [51754257.0], [51754431.0], [51754669.0], [51754669.0], [51754706.0], [51754761.0], [51754954.0], [51754954.0], [51754966.0], [51754966.0], [51754966.0], [51754966.0], [51754966.0], [51755450.0], [51755450.0], [51755450.0], [51755456.0], [51755773.0], [51755773.0], [51756440.0], [51756440.0], [51756628.0], [51756628.0], [51757005.0], [51757005.0], [51757816.0], [51757816.0], [51757816.0], [51757816.0], [51757816.0], [51757816.0], [51758009.0], [51758009.0], [51758009.0], [51758291.0], [51758291.0], [51758694.0], [51758798.0], [51758798.0], [51759166.0], [51759614.0], [51759922.0], [51760069.0], [51760818.0], [51760869.0], [51761042.0], [51761064.0], [51761398.0], [51761529.0], [51761965.0], [51761965.0], [51762219.0], [51762621.0], [51762621.0], [51762865.0], [51763071.0], [51763071.0], [51763084.0], [51763317.0], [51763432.0], [51763533.0], [51763660.0], [51763951.0], [51763951.0], [51764010.0], [51764200.0], [51764200.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764599.0], [51764848.0], [51764848.0], [51764848.0], [51764848.0], [51764893.0], [51764990.0], [51765182.0], [51765271.0], [51766401.0], [51766401.0], [51766401.0], [51766401.0], [51766401.0], [51766401.0], [51766401.0], [51766537.0], [51766537.0], [51766537.0], [51767195.0], [51767195.0], [51767195.0], [51767664.0], [51767858.0], [51767858.0], [51768221.0], [51769381.0], [51770130.0], [51770130.0], [51771228.0], [51771228.0], [51771228.0], [51771228.0], [51771813.0], [51771813.0], [51771813.0], [51771813.0], [51773116.0], [51773703.0], [51774690.0], [51774690.0], [51774690.0], [51774745.0], [51774745.0], [51776075.0], [51776075.0], [51776075.0], [51776075.0], [51776075.0], [51776075.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51776404.0], [51777153.0], [51777153.0], [51777153.0], [51767195.0], [51767195.0], [51767195.0], [51767195.0], [51767195.0], [51767195.0], [51767195.0], [51767195.0], [51771676.0], [51767195.0], [51764399.0], [51775655.0], [51775655.0], [51774011.0], [51774011.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51767837.0], [51773396.0], [51775059.0], [51775059.0], [51775059.0], [51767060.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776124.0], [51776420.0], [51776420.0], [51776420.0], [51776420.0], [51776420.0], [51776431.0], [51768507.0], [51775636.0], [51775636.0], [51775636.0], [51775636.0], [51776142.0], [51776142.0], [51776142.0], [51776142.0], [51776142.0], [51776142.0], [51775554.0], [51775554.0], [51775554.0], [51775554.0], [51775554.0], [51775554.0], [51775554.0], [51775554.0], [51775554.0], [51775554.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51774745.0], [51767469.0], [51768779.0], [51769221.0], [51771448.0], [51773985.0], [51775664.0], [51775664.0], [51775664.0], [51775664.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51768797.0], [51768797.0], [51768797.0], [51770699.0], [51770699.0], [51770699.0], [51770699.0], [51770699.0], [51772683.0], [51772683.0], [51773571.0], [51773571.0], [51773571.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51774422.0], [51776581.0], [51776581.0], [51776581.0], [51776581.0], [51768607.0], [51770269.0], [51773473.0], [51773473.0], [51773473.0], [51773473.0], [51773473.0], [51774146.0], [51772049.0], [51772049.0], [51773985.0], [51773985.0], [51767469.0], [51769221.0], [51773985.0], [51775017.0], [51775664.0], [51775664.0], [51775664.0], [51775664.0], [51775664.0], [51775664.0], [51775664.0], [51775664.0], [51774837.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51759910.0], [51775059.0], [51775059.0], [51775059.0], [51771448.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51773473.0], [51775581.0], [51767469.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51772045.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51772094.0], [51773238.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51762323.0], [51771769.0], [51771769.0], [51771769.0], [51771769.0], [51770269.0], [51773985.0], [51774452.0], [51774452.0], [51775581.0], [51775832.0], [51775832.0], [51775832.0], [51775832.0], [51775832.0], [51775832.0], [51775832.0], [51775832.0], [51775832.0], [51775832.0], [51775832.0], [51747313.0], [51776581.0], [51776581.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747268.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51747313.0], [51772049.0], [51776041.0], [51776041.0], [51776041.0], [51776041.0], [51776041.0], [51772610.0], [51772610.0], [51768507.0], [51772610.0], [51776041.0], [51774837.0], [51774837.0], [51774837.0], [51774837.0], [51774837.0], [51774837.0], [51776041.0], [51773615.0], [51774837.0], [51774837.0], [51774837.0], [51774837.0], [51774837.0], [51774837.0], [51776041.0], [51776041.0], [51776041.0], [51776041.0]]
            #Item = [[4010.0], [50020.0], [20020.0], [20030.0], [21010.0], [21020.0], [21030.0], [1010.0], [3010.0], [11010.0], [11020.0], [11030.0], [12010.0], [12020.0], [12030.0], [1010.0], [13505.0], [14505.0], [7015.0], [7025.0], [7515.0], [15010.0], [16010.0], [1015.0], [1030.0], [6015.0], [6510.0], [9010.0], [6010.0], [9010.0], [10010.0], [2010.0], [3010.0], [3010.0], [12010.0], [2010.0], [1010.0], [12010.0], [13010.0], [1010.0], [1010.0], [1010.0], [3010.0], [3020.0], [4010.0], [1010.0], [1010.0], [7030.0], [7070.0], [1010.0], [2010.0], [3010.0], [1010.0], [1010.0], [2010.0], [12510.0], [19010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [2010.0], [3010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [2020.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [2010.0], [3010.0], [4010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [3010.0], [1010.0], [1010.0], [2010.0], [3010.0], [4010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [2010.0], [3010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [2010.0], [1010.0], [3010.0], [4010.0], [7410.0], [8410.0], [8420.0], [24010.0], [24020.0], [1010.0], [1010.0], [1010.0], [2010.0], [8010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [3010.0], [5010.0], [7010.0], [9010.0], [4010.0], [1010.0], [1020.0], [1010.0], [2010.0], [3010.0], [1010.0], [1020.0], [1030.0], [2010.0], [2020.0], [3010.0], [3020.0], [3030.0], [24010.0], [25010.0], [1010.0], [2010.0], [1010.0], [2010.0], [3010.0], [4010.0], [1010.0], [12010.0], [13010.0], [18010.0], [19010.0], [20010.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [1060.0], [2010.0], [2020.0], [2030.0], [2040.0], [2050.0], [2060.0], [1010.0], [1020.0], [1040.0], [1050.0], [1060.0], [1080.0], [10010.0], [10020.0], [6040.0], [9040.0], [1010.0], [1010.0], [1010.0], [1020.0], [1010.0], [2010.0], [1010.0], [2010.0], [9030.0], [27010.0], [6020.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [1060.0], [1070.0], [1080.0], [2060.0], [1010.0], [2010.0], [1010.0], [2010.0], [3010.0], [2015.0], [2525.0], [4010.0], [1010.0], [3010.0], [4010.0], [3020.0], [5010.0], [8020.0], [1010.0], [1010.0], [1010.0], [8010.0], [9010.0], [20010.0], [2030.0], [6010.0], [6020.0], [6030.0], [6040.0], [6050.0], [6060.0], [54010.0], [55010.0], [55020.0], [56010.0], [57010.0], [1020.0], [1030.0], [2010.0], [2020.0], [3010.0], [3020.0], [4010.0], [4020.0], [5010.0], [5020.0], [6010.0], [6020.0], [1010.0], [1010.0], [2010.0], [1010.0], [2010.0], [7030.0], [1010.0], [2010.0], [1010.0], [4010.0], [4020.0], [5010.0], [5020.0], [5030.0], [5040.0], [29010.0], [1010.0], [1010.0], [1010.0], [1010.0], [2010.0], [3010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [4020.0], [5010.0], [6010.0], [6020.0], [1010.0], [1010.0], [2010.0], [1010.0], [6030.0], [9030.0], [2010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [2010.0], [3010.0], [1010.0], [1010.0], [1010.0], [2010.0], [16010.0], [17010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [2010.0], [6010.0], [7010.0], [9010.0], [10010.0], [11010.0], [13010.0], [13020.0], [13030.0], [1010.0], [2010.0], [3010.0], [4010.0], [5010.0], [6010.0], [7010.0], [8010.0], [9010.0], [10010.0], [11010.0], [12010.0], [13010.0], [14010.0], [15010.0], [16010.0], [17010.0], [18010.0], [19010.0], [20010.0], [24010.0], [24020.0], [24030.0], [24040.0], [25060.0], [25070.0], [25080.0], [6010.0], [7010.0], [1010.0], [1020.0], [2010.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [1060.0], [5040.0], [8040.0], [1010.0], [10020.0], [3010.0], [4010.0], [10010.0], [10020.0], [10030.0], [10040.0], [10060.0], [11010.0], [11020.0], [11030.0], [11040.0], [11060.0], [17010.0], [12010.0], [13010.0], [14010.0], [14030.0], [14040.0], [1010.0], [2010.0], [3010.0], [4010.0], [5010.0], [6010.0], [7010.0], [8010.0], [9010.0], [7010.0], [8010.0], [8020.0], [9010.0], [10010.0], [9010.0], [9020.0], [9030.0], [10020.0], [10030.0], [11020.0], [12010.0], [3010.0], [17010.0], [17020.0], [17030.0], [17040.0], [18010.0], [18020.0], [18030.0], [18040.0], [18050.0], [5010.0], [1010.0], [2010.0], [3010.0], [3020.0], [3030.0], [3020.0], [3060.0], [7010.0], [7020.0], [7030.0], [7040.0], [7050.0], [7060.0], [1010.0], [9010.0], [10010.0], [11010.0], [12010.0], [12020.0], [13010.0], [14010.0], [14020.0], [15010.0], [16010.0], [16020.0], [17010.0], [18010.0], [18020.0], [19010.0], [20010.0], [20020.0], [20030.0], [20040.0], [20050.0], [1010.0], [3010.0], [6010.0], [1010.0], [2010.0], [2010.0], [3010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [2010.0], [3010.0], [1010.0], [2010.0], [1010.0], [1010.0], [28010.0], [29010.0], [6010.0], [27010.0], [6010.0], [27010.0], [2010.0], [9010.0], [30010.0], [7010.0], [28010.0], [1010.0], [2010.0], [5010.0], [26010.0], [1010.0], [11010.0], [33010.0], [6010.0], [27010.0], [10010.0], [31010.0], [18010.0], [40010.0], [1010.0], [2010.0], [1010.0], [13010.0], [34010.0], [14010.0], [36010.0], [15010.0], [35010.0], [17010.0], [38010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [22010.0], [5070.0], [1010.0], [2010.0], [1010.0], [1010.0], [5010.0], [5020.0], [5030.0], [5040.0], [5050.0], [6010.0], [6020.0], [6030.0], [6040.0], [9010.0], [10010.0], [1010.0], [20010.0], [40010.0], [40010.0], [41010.0], [1010.0], [21010.0], [42010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [2010.0], [3010.0], [3020.0], [3030.0], [5010.0], [5020.0], [1010.0], [2010.0], [3010.0], [1010.0], [7040.0], [7050.0], [1010.0], [2010.0], [25010.0], [25020.0], [1010.0], [2010.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [1060.0], [11010.0], [12010.0], [13010.0], [1010.0], [2010.0], [1010.0], [2505.0], [3505.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [6010.0], [6020.0], [1010.0], [1010.0], [2010.0], [1010.0], [1020.0], [2010.0], [3010.0], [4010.0], [10010.0], [11010.0], [12010.0], [14010.0], [15010.0], [16010.0], [17010.0], [2010.0], [2020.0], [2030.0], [2040.0], [1010.0], [1010.0], [1010.0], [1010.0], [1010.0], [2010.0], [2020.0], [2030.0], [3010.0], [3020.0], [4020.0], [1010.0], [1020.0], [1030.0], [3045.0], [3535.0], [16035.0], [1010.0], [1010.0], [2010.0], [1010.0], [1010.0], [1010.0], [3010.0], [4010.0], [4020.0], [7010.0], [7020.0], [1010.0], [2010.0], [3010.0], [4010.0], [1010.0], [1010.0], [1010.0], [2010.0], [3010.0], [9010.0], [10010.0], [1010.0], [1020.0], [2010.0], [3010.0], [4010.0], [5010.0], [2010.0], [2020.0], [3010.0], [3020.0], [3030.0], [3040.0], [3050.0], [4050.0], [5030.0], [5040.0], [6010.0], [6020.0], [6030.0], [6040.0], [8010.0], [8040.0], [1010.0], [2010.0], [3010.0], [16010.0], [16020.0], [16040.0], [3010.0], [3020.0], [3030.0], [3050.0], [3060.0], [2010.0], [15010.0], [22010.0], [1010.0], [2010.0], [9010.0], [9020.0], [1010.0], [1020.0], [2010.0], [2020.0], [3010.0], [3020.0], [3030.0], [4010.0], [4020.0], [4030.0], [4040.0], [5010.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [1060.0], [1070.0], [1080.0], [10010.0], [10020.0], [10030.0], [10040.0], [1010.0], [1010.0], [1020.0], [1030.0], [4010.0], [5020.0], [5030.0], [5040.0], [5050.0], [5060.0], [5070.0], [6010.0], [6020.0], [7010.0], [8010.0], [8020.0], [9010.0], [9020.0], [9030.0], [9040.0], [9050.0], [9060.0], [10010.0], [10020.0], [10030.0], [10040.0], [11010.0], [11020.0], [1010.0], [1020.0], [1030.0], [2010.0], [2020.0], [1010.0], [5010.0], [11010.0], [12010.0], [13010.0], [14010.0], [9010.0], [10010.0], [10020.0], [10030.0], [10040.0], [11010.0], [1010.0], [2010.0], [3010.0], [4010.0], [5010.0], [12010.0], [12020.0], [13010.0], [13020.0], [13030.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [1060.0], [2010.0], [2020.0], [2030.0], [2040.0], [2050.0], [2060.0], [3010.0], [3020.0], [3030.0], [3040.0], [3050.0], [3060.0], [4010.0], [4020.0], [4030.0], [4040.0], [4050.0], [4060.0], [5010.0], [5020.0], [5030.0], [5040.0], [5050.0], [6010.0], [6020.0], [6030.0], [6040.0], [1010.0], [4030.0], [1010.0], [5030.0], [10030.0], [2010.0], [2020.0], [2030.0], [3010.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [1060.0], [1070.0], [1080.0], [2010.0], [2020.0], [2030.0], [2040.0], [2050.0], [2060.0], [2070.0], [2080.0], [2010.0], [3010.0], [22040.0], [1010.0], [1020.0], [1030.0], [1040.0], [1050.0], [2010.0], [3010.0], [1010.0], [2010.0], [3010.0], [7010.0], [7020.0], [7030.0], [7040.0], [8010.0], [8020.0], [8030.0], [8040.0], [8050.0], [8060.0], [8070.0], [8080.0], [1010.0], [2010.0], [3010.0], [6010.0], [16010.0], [3010.0], [12010.0], [12020.0], [12030.0], [12040.0], [12050.0], [1010.0], [14010.0], [15010.0], [5010.0], [8010.0], [3010.0], [7010.0], [6010.0], [5010.0], [7010.0], [7020.0], [7030.0], [7040.0], [7050.0], [8010.0], [8020.0], [8030.0], [11010.0], [6010.0], [6020.0], [6030.0], [6040.0], [6050.0], [6060.0], [6070.0], [6080.0], [7010.0], [7020.0], [7030.0], [7040.0], [7050.0], [7060.0], [7070.0], [7080.0], [8010.0], [8020.0], [8030.0], [8040.0], [2010.0], [2020.0], [2030.0], [2010.0], [1010.0], [2010.0], [2020.0], [2030.0], [2040.0], [2050.0], [2060.0], [3010.0], [4010.0], [5010.0], [6010.0], [7010.0], [1010.0], [1020.0], [1030.0], [2010.0], [3010.0], [3020.0], [4010.0], [4020.0], [4030.0], [4040.0], [4050.0], [4060.0], [5010.0], [5020.0], [5030.0], [5040.0], [5050.0], [5060.0], [5070.0], [4010.0], [1010.0], [14020.0], [7020.0], [7030.0], [7040.0], [7050.0], [7060.0], [6010.0], [6020.0], [6030.0], [6040.0], [6050.0], [7010.0], [7020.0], [8010.0], [8020.0], [8030.0], [8040.0], [8050.0], [9010.0], [9020.0], [9030.0], [9040.0], [9050.0], [9060.0], [9070.0], [8010.0], [8020.0], [8030.0], [8040.0], [8050.0], [5010.0], [5010.0], [1010.0], [1020.0], [1030.0], [1040.0], [5010.0], [6010.0], [7010.0], [7020.0], [7030.0], [8010.0], [8020.0], [9010.0], [9020.0], [9030.0], [9040.0], [9050.0], [9060.0], [9070.0], [9080.0], [10010.0], [10020.0], [10030.0], [10040.0], [10050.0], [10060.0], [10070.0], [10080.0], [11010.0], [11020.0], [11030.0], [11040.0], [11050.0], [11060.0], [11070.0], [11080.0], [8060.0], [9010.0], [10010.0], [10020.0], [15010.0], [19010.0], [1010.0], [2010.0], [5010.0], [11010.0], [11020.0], [11030.0], [11040.0], [11050.0], [13010.0], [16010.0], [16020.0], [16030.0], [16040.0], [16050.0], [1010.0], [4010.0], [5010.0], [1010.0], [2010.0], [3010.0], [3020.0], [3030.0], [3040.0], [3050.0], [3060.0], [3070.0], [4010.0], [5010.0], [6010.0], [6020.0], [6030.0], [6040.0], [6050.0], [6060.0], [6070.0], [7010.0], [8010.0], [9010.0], [9020.0], [9030.0], [9040.0], [9050.0], [9060.0], [9070.0], [10010.0], [11010.0], [11020.0], [1020.0], [1030.0], [1040.0], [1050.0], [2010.0], [2020.0], [3010.0], [3020.0], [4010.0], [4020.0], [4030.0], [4040.0], [4050.0], [4060.0], [4070.0], [5010.0], [5020.0], [5030.0], [6010.0], [6020.0], [6030.0], [7010.0], [7020.0], [8010.0], [8020.0], [8030.0], [8040.0], [8050.0], [9010.0], [9020.0], [9030.0], [9040.0], [9050.0], [9060.0], [9070.0], [29010.0], [13010.0], [13020.0], [13030.0], [13040.0], [13050.0], [4010.0], [6010.0], [1010.0], [5010.0], [13060.0], [6010.0], [6020.0], [6030.0], [7010.0], [7020.0], [7030.0], [3010.0], [2010.0], [10010.0], [10020.0], [10030.0], [10040.0], [10050.0], [10060.0], [4010.0], [5010.0], [6010.0], [7010.0]]

            #for i in range(1,1176):
             #   self.dbCursor.execute("""UPDATE orders.root Set QTE_Release_Status = ? where sales_order = ? AND item = ?""", str('Pending'),str(int(SO[i-1][0])),str(int(Item[i-1][0])))
              #  self.conn.commit()



        def OnBtnRemoveHoliday(self,event = None):
            if self.m_DateRemoveHoliday.GetValue():
                currHolidaylist = 'SELECT * from Holidays'
                holidays = self.dbCursor.execute(currHolidaylist).fetchall()
                holidays_list = []
                for holiday in holidays:
                    for date in holiday:
                        holidays_list.append(date)
                removedate = str(self.m_DateRemoveHoliday.GetValue())
                removedate1 = datetime.datetime.strptime(removedate, "%d/%m/%Y %H:%M:%S").strftime("%m/%d/%Y %H:%M:%S")
                if removedate1 in holidays_list:
                    sqlRemove = 'Delete top(1) from Holidays where Holidays = \'{}\''.format(str(removedate1))
                    self.dbCursor.execute(sqlRemove)
                    self.conn.commit()
                    msgbox = wx.MessageBox('Holiday removed successfully', 'Alert')
                    #print removedate1
                else:
                    msgbox = wx.MessageBox('This Holiday is not added in database, you cannot remove holiday which is not added in database. Click Display Holiday List button to view which all Holidays are added in the database.', 'Alert')





