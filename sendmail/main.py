#!/usr/bin/env python
#coding=utf-8

import wx,shelve,dumbdbm,anydbm,time,thread
import wx.lib.filebrowsebutton
import wx.lib.scrolledpanel

from smtpProcessor.smtp import TestSmtpServer, Smtp
from xlProcessor.xprocessor import XProcessor

#发送状态窗口
class SendStatusPanel(wx.lib.scrolledpanel.ScrolledPanel):
    def __init__(self,parent):
        print "aaaaaaa"
        wx.lib.scrolledpanel.ScrolledPanel.__init__(self, parent, -1)
        
        user_name = wx.StaticText(self, -1, u'姓名')
        email_address = wx.StaticText(self, -1, u'邮箱')
        status = wx.StaticText(self, -1, u'状态')
        #print status.GetId().__class__
        #设置sizer
        self.sizer = wx.FlexGridSizer(0,3,10,10)
        self.sizer.Add(user_name,0,0)
        self.sizer.Add(email_address,0,0)
        self.sizer.Add(status,0,0)
        self.SetSizer(self.sizer)#把sizer与框架关联起来
        self.FitInside()
        self.SetAutoLayout(1)
        self.SetupScrolling()
        
    #添加提示全部发送成功的静态文本
    def addSuccessNoticeToSizer(self):
        success = wx.StaticText(self, -1, u'恭喜，已全部发送成功')
        self.sizer.Add(success,0,0)
    #添加单个部件到Sizer
    def addSingleContextToSizer(self,user_name,email,status_string):
            user_name = wx.StaticText(self, -1, unicode(user_name))
            email_address = wx.StaticText(self, -1, unicode(email))
            status = wx.StaticText(self, -1, unicode(status_string))
            self.sizer.Add(user_name,0,0)
            self.sizer.Add(email_address,0,0)
            self.sizer.Add(status,0,0)
    #添加多个部件到Sizer
    def addContextToSizer(self,emailsAndContents):
        i = 1
        for sheetname in emailsAndContents.keys():
            cls = emailsAndContents.get(sheetname, [])
            for cl in cls:
                user_name = wx.StaticText(self, -1, unicode(cl.username))
                email_address = wx.StaticText(self, -1, unicode(cl.email))
                status = wx.StaticText(self, i , u'等待发送')
                status.SetId(i)
                status.SetBackgroundColour(u"red")
                self.sizer.Add(user_name,0,0)
                self.sizer.Add(email_address,0,0)
                self.sizer.Add(status,0,0)
                i += 1
        
#发送邮件面板
class SendEmailPanel(wx.Panel):
    def __init__(self,parent):
        wx.Panel.__init__(self,parent,pos=(0,0),size = (360,260))
        
        #账号显示
        self.display_smpt_account = wx.StaticBox(id=-1,
              label=u'发送邮件账号', name=u'smtpAccount',
              parent=self, pos=wx.Point(8, 10), size=wx.Size(320,80),
              style=0)
        wx.StaticText(self, -1, u'邮箱:',pos = (17,35))
        self.your_send_eamil = wx.StaticText(self, -1, u'your email',pos=(95,33),
                                size=(175, -1))
        wx.StaticText(self, -1, u'STMP:',pos = (17,55))
        self.smtp_server_text = wx.StaticText(self, -1, u"your smtp server", pos = (95,55),size=(175, -1))
        
        
        #选择excel文件
        self.choose_excel_file = wx.StaticBox(id=-1,
              label=u'选择excel文件', name=u'chooseExcelFile',
              parent=self, pos=wx.Point(8, 100), size=wx.Size(320,80),
              style=0)
        self.is_show_status = False
        #源图片文件夹
        self.choose_excel_path = wx.lib.filebrowsebutton.FileBrowseButtonWithHistory(buttonText=u'浏览',
              dialogTitle=u'选择excel文件',
              id=-1,
              labelText='',
              parent=self, pos=wx.Point(17, 125),
              size=wx.Size(300, 40), startDirectory='.',
              toolTip=u"选择excel文件",
              changeCallback=self._setExcelFilePath)
        self.choose_excel_path.SetName(u'chooseExcelPath')
        
        wx.StaticText(self, -1, u'(支持excel的文件后缀有xlsx/xlsm/xltx/xltm, 不再支持xls)',pos = (10,165))
        
        #确认发送按钮
        self.send_email_button = wx.Button(id=-1,
              label=u'确认文件并开始发送邮件',
              name=u'comfirmAndSendEmail', parent=self, pos=wx.Point(90,182),
              size=wx.Size(160, 47), style=0)
        self.send_email_button.Bind(wx.EVT_BUTTON, self.onSendEmailButton,
              id=-1)
        self.excel_file_path = ''
        self.getEmailAndSmtpServer()
        
        self.status_frame = None
        
    #读取邮箱，smtp服务器
    def getEmailAndSmtpServer(self):
        data = shelve.open(u'data')
        if not data :
            return
        username = data.get(u'username','')
        smtp_server_name = data.get(u'smtp_server_name','')
        data.close()
        self.your_send_eamil.Label = unicode(username)
        self.smtp_server_text.Label = unicode(smtp_server_name)
        
    #self.choose_excel_path的回调函数
    def _setExcelFilePath(self,event):
        if self.is_show_status:
            return
        self.is_show_status = True
        print "bbbbbbb"
        self.excel_file_path = unicode(self.choose_excel_path.GetValue())
        self.status_frame = ''
        if not self.excel_file_path:
            return
        
        if self.status_frame:
            self.status_frame.Close()
            self.status_frame.Destroy()
            self.status_frame = None
        try:
            xp = XProcessor()
            xp.start(self.excel_file_path)
            self.emailsAndContents = xp.content
        except Exception, e:
            notice = wx.MessageDialog(self, (e.msg if hasattr(e, "msg") else unicode(e)),caption=u'警告',
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
            notice.ShowModal()
            notice.Destroy()
            self.is_show_status = False
        else:
            self.status_frame = wx.Frame(self,-1,u"发送状态--未发送",size=(360,-1),pos=(815,200))
            self.status_frame.SetIcon(wx.Icon(name='email.ico', type=wx.BITMAP_TYPE_ICO))
            self.send_status_frame = SendStatusPanel(self.status_frame)
            self.send_status_frame.addContextToSizer(self.emailsAndContents)
            self.status_frame.Show()
            self.is_show_status = False
        
    #响应确认文件并开始发送邮件按钮
    def onSendEmailButton(self,event):
        if not self.excel_file_path:
            return
        thread.start_new_thread(self.startSendSalaryEmail,())
    
    #开始发送邮件
    def startSendSalaryEmail(self):    
        try:
            #print self.emailsAndContents.to_emailList
            data = shelve.open(u'data')
            if not data :
                return
            username = data.get(u'username','')
            password = data.get(u'password','')
            smtp_server_name = data.get(u'smtp_server_name','')
            port = int(data.get(u'port',25))
            ssl_type = int(data.get(u'ssl_type',0))
            data.close()
            self.send_server = Smtp(username,password,smtp_server_name,port, ssl_type)
            self.send_server.startSmtpServer()
            
            #状态窗口
            #self.status_frame = wx.Frame(self,-1,u"发送状态",size=(360,-1))
            #self.status_frame.SetIcon(wx.Icon(name='email.ico', type=wx.BITMAP_TYPE_ICO))
            #self.send_status_frame = SendStatusPanel(self.status_frame)
            #self.status_frame.Show()
            
            #发送
            sended_error_email = []
            i = 1
            cls = []
            for sheetname in self.emailsAndContents.keys():
                cls += self.emailsAndContents.get(sheetname, [])
            for cl in cls:
                try:
                    self.send_server.sendEmail(cl.email, username, u'工资条', cl.content)
                except Exception:
                    #查看网络是否断开
                    test = TestSmtpServer()
                    if thread.start_new_thread(test.findSmtpServer,(smtp_server_name, port, ssl_type)):
                        #print test.find_smtp_server(smtp_server_name, port)
                        notice = wx.MessageDialog(self,u'互联网连接断开，请重新连接后,单击“是”以确定继续发送邮件，否则单击“否”则停止发送。',caption=u'警告',
                                    style=wx.YES_NO|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
                        is_yes = notice.ShowModal()
                        
                        if is_yes == wx.ID_YES:
                            if test.findSmtpServer(smtp_server_name, port, ssl_type):
                                sended_error_email.append((cl.email, cl.username))
                                self.send_status_frame.FindWindowById(i).SetBackgroundColour(u"blue")
                                self.send_status_frame.FindWindowById(i).Label = u'发送失败'
                                i +=1
                                return
                            self.send_server.startSmtpServer()
                            self.send_server.sendEmail(cl.email, username, u'工资条', cl.content)
                            self.send_status_frame.FindWindowById(i).SetBackgroundColour(u"green")
                            self.send_status_frame.FindWindowById(i).Label = u'发送成功'
                            i +=1
                            continue
                            print notice.ShowModal
                        else :
                            sended_error_email.append((cl.email, cl.username))
                            self.send_status_frame.FindWindowById(i).SetBackgroundColour(u"blue")
                            self.send_status_frame.FindWindowById(i).Label = u'发送失败'
                            i +=1
                            return
                        notice.Destroy()
                    #其他错误暂没处理
                    sended_error_email.append((cl.email, cl.username))
                    self.send_status_frame.FindWindowById(i).SetBackgroundColour(u"blue")
                    self.send_status_frame.FindWindowById(i).Label = u'发送失败'
                    i +=1
                    #sended_error_email.append(email)
                    #self.send_status_frame.FindWindowById(i).SetBackgroundColour(u"blue")
                    #self.send_status_frame.FindWindowById(i).Label = u'发送失败'
                    #i +=1
                    #self.sendEmailAgainIfError(username,password,smtp_server_name,port,
                            #email,self.emailsAndContents.email_content[email])
                        
                else:
                    self.send_status_frame.FindWindowById(i).SetBackgroundColour(u"green")
                    self.send_status_frame.FindWindowById(i).Label = u'发送成功'
                    i +=1
            self.send_server.startSmtpServer()
        except Exception, e:
            print e
            notice = wx.MessageDialog(self,u'请确保互联网是否连接。若已连接请确保邮箱和密码正确。你可以尝试重新设置发送服务器以解决该问题。',caption=u'登录服务器失败',
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
            notice.ShowModal()
            notice.Destroy()
            return
        else:
            if not sended_error_email:
                notice = wx.MessageDialog(self,u'恭喜，已全部发送成功！',
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
                self.status_frame.SetTitle(u"发送状态--已全部发送成功")
                
                notice.ShowModal()
                notice.Destroy()
            else:
                message = u'以下是发送错误的姓名和邮箱：'
                for email, username in sended_error_email:
                    string = unicode(u'\n'+username+u'，'+email)
                    message += string
                notice = wx.MessageDialog(self,message,
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
                notice.ShowModal()
                notice.Destroy()
                self.status_frame.SetTitle(u"发送状态--存在发送错误")

#设置服务器面板
class SetSmtpServerPanel(wx.Panel):
    def __init__(self,parent):
        wx.Panel.__init__(self,parent,pos=(0,0),size = (360,260))
        
        self.set_smtp_server = wx.StaticBox(id=-1,
              label=u'设置发送邮件服务器', name=u'setSmtpSever',
              parent=self, pos=wx.Point(8, 10), size=wx.Size(320, 180),
              style=0)
        #邮箱
        wx.StaticText(self, -1, u'邮箱:',pos = (17,35))
        self.eamil_text = wx.TextCtrl(self, -1, u'',pos=(135,33),
                                size=(175, -1))
        #密码
        wx.StaticText(self, -1, u'密码:',pos = (17,75))
        self.password_text = wx.TextCtrl(self, -1, u"", pos = (135,73),size=(175, -1),
                              style=wx.TE_PASSWORD)
        #smtp服务器
        wx.StaticText(self, -1, u'SMTP服务器:',pos = (17,115))
        self.smtp_server_text = wx.TextCtrl(self, -1, u"", pos = (135,113),size=(175, -1))
        
        #smtp服务器端口
        wx.StaticText(self, -1, u'SMTP服务器端口:',pos = (17,155))
        self.port_text = wx.TextCtrl(self, -1, u"", pos = (135,153),size=(60, -1))
        
        
        #smtp端口
        wx.StaticText(self, -1, u'加密方法:',pos = (17,195))
        self.no_ssl_text = wx.RadioButton(id=-1,
              label=u'无', name=u'no_ssl', parent=self,
              pos=wx.Point(140, 195), size=wx.Size(48, 14), style=0)
        self.no_ssl_text.SetValue(True)
        self.ssl_text = wx.RadioButton(id=-1,
              label=u'SSL', name=u'ssl', parent=self,
              pos=wx.Point(190, 195), size=wx.Size(48, 14), style=0)
        self.tls_text = wx.RadioButton(id=-1,
              label=u'TLS', name=u'tls', parent=self,
              pos=wx.Point(242, 195), size=wx.Size(48, 14), style=0)
        
        #确定并测试按钮
        self.comfirm_and_test_smtp_button = wx.Button(id=-1,
              label=u'确认并测试',
              name=u'comfirmAndTest', parent=self, pos=wx.Point(235,232),
              size=wx.Size(96, 32), style=0)
        self.comfirm_and_test_smtp_button.Bind(wx.EVT_BUTTON, self.onComfirmAndTestSmtpButton,
              id=-1)
        self.testing = False
        self.getInfo()
        
    #读取存储的邮箱，密码，smtp服务器
    def getInfo(self):
        data = shelve.open(u'data')
        if not data :
            return
        username = data.get(u'username','')
        password = data.get(u'password','')
        smtp_server_name = data.get(u'smtp_server_name','')
        port = int(data.get(u'port',25))
        ssl_type = int(data.get(u'ssl_type', 0))
        if ssl_type == 2:
            self.tls_text.SetValue(True)
        elif ssl_type == 1:
            self.ssl_text.SetValue(True)
        else:
            self.no_ssl_text.SetValue(True)
        self.eamil_text.SetValue(unicode(username))
        self.password_text.SetValue(unicode(password))
        self.smtp_server_text.SetValue(unicode(smtp_server_name))
        self.port_text.SetValue(unicode(port))
        data.close()
    
    #响应确认按钮    
    def onComfirmAndTestSmtpButton(self,event):
        if self.testing:
            return
        username = unicode(self.eamil_text.GetValue())
        password = unicode(self.password_text.GetValue())
        smtp_server_name = unicode(self.smtp_server_text.GetValue())
        port = int(self.port_text.GetValue() if self.port_text.GetValue() else 25)
        ssl_type = 0
        if self.ssl_text.GetValue():
            ssl_type = 1
        elif self.tls_text.GetValue():
            ssl_type = 2
        
        if not username or not password or not smtp_server_name:
            notice = wx.MessageDialog(self,u'邮箱，密码，STMP服务器，端口号都不能为空!',caption=u'警告',
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
            notice.ShowModal()
            notice.Destroy()
            return 
        
        #测试smtp服务器
        
        self.testing = True
        self.comfirm_and_test_smtp_button.SetLabel(u"正在测试...")
        def TestServer():
            testSmtpServer = TestSmtpServer()
            string = testSmtpServer.findSmtpServer(smtp_server_name, port, ssl_type)
            if string :
                notice = wx.MessageDialog(self,unicode(string),caption=u'警告',
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
                notice.ShowModal()
                notice.Destroy()
                self.comfirm_and_test_smtp_button.SetLabel(u"确定并测试")
                self.testing = False
                return 
            string = testSmtpServer.loginTest(username, password)
            if string :
                notice = wx.MessageDialog(self,unicode(string),caption=u'警告',
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
                notice.ShowModal()
                notice.Destroy()
                self.comfirm_and_test_smtp_button.SetLabel(u"确定并测试")
                self.testing = False
                return 
        
            notice = wx.MessageDialog(self,unicode(u'连接服务器成功！现在你可以发送邮件了！'),caption=u'测试成功',
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
            notice.ShowModal()
            notice.Destroy()
        
            #存储邮箱，密码，smtp服务器
            data = shelve.open(u'data')
            data['username'] = username
            data['password'] = password
            data['smtp_server_name'] = smtp_server_name
            data['port'] = port
            data['ssl_type'] = ssl_type
            data.close()
            self.comfirm_and_test_smtp_button.SetLabel(u"确定并测试")
            self.testing = False
            
        thread.start_new_thread(TestServer, ())
#主窗口
class MainFrame(wx.Frame):
    def __init__(self, parent=None, id=-1, title=''):
        wx.Frame.__init__(self, parent, id, title,size = (360,344),pos = (450,200))
        self.SetIcon(wx.Icon(name='email.ico', type=wx.BITMAP_TYPE_ICO))
        #设置工具栏
        self.CreateStatusBar() 
        toolbar = self.CreateToolBar(style=(wx.TB_HORZ_LAYOUT|wx.TB_NOICONS | wx.TB_TEXT | wx.TB_3DBUTTONS)) 
        qtool_send = toolbar.AddLabelTool(wx.ID_ANY,u'开始发送邮件',wx.Bitmap("toorbar.png"))
        toolbar.AddSeparator()
        qtool2_set_smtp = toolbar.AddLabelTool(wx.ID_ANY,u'设置发送邮件服务器',wx.Bitmap("toorbar.png")) 
        toolbar.AddSeparator()
        toolbar.Realize() 
        self.Bind(wx.EVT_MENU, self.showSendEmailWindow, qtool_send)
        self.Bind(wx.EVT_MENU, self.showSetSmtpServerWindow,qtool2_set_smtp)
        
        #是否创建画板标志
        self.createdSetSmtpServerWindowFlag = False#是否生成set_panel的内容
        self.createdSendEmailWindowFlag = False#是否生成send_email_panel内容
        
        #是否是第一次打开
        data = shelve.open(u'data')
        if not data:
            notice = wx.MessageDialog(self,unicode(u'首次打开请设置发送服务器。'),caption=u'提示',
                                    style=wx.OK|wx.ICON_INFORMATION,
                                    pos=wx.DefaultPosition)
            notice.ShowModal()
            notice.Destroy()
            self.createSetSmtpServerWindow()
        else:
            self.createSendEmailWindow()
        
        
    
    #生成发送邮件窗口
    def createSendEmailWindow(self):
        #设置发送邮件面板
        self.send_email_panel = SendEmailPanel(self)
        self.createdSendEmailWindowFlag = True
        
        
    #生成设置窗口
    def createSetSmtpServerWindow(self):
        #设置发送服务器面板
        self.set_panel = SetSmtpServerPanel(self)
        self.createdSetSmtpServerWindowFlag = True
    
    
    u'''以下事件处理'''
        
    
    def showSetSmtpServerWindow(self,event):
        if not self.createdSendEmailWindowFlag:
            self.createSendEmailWindow()
        self.send_email_panel.Hide()
        
        if not self.createdSetSmtpServerWindowFlag:
            self.createSetSmtpServerWindow()
        else:
            self.set_panel.Show()
            
    def showSendEmailWindow(self,event):
        if not self.createdSetSmtpServerWindowFlag:
            self.createSetSmtpServerWindow()
        self.set_panel.Hide()
        
        if self.createdSendEmailWindowFlag:
            self.createSendEmailWindow()
        else:
            self.send_email_panel.Show()
        self.send_email_panel.getEmailAndSmtpServer()


if __name__ == u"__main__":
    app = wx.App(redirect=False)
    frame = MainFrame(None,-1,u'发送工资邮件小工具')
    frame.Show()
    app.MainLoop()