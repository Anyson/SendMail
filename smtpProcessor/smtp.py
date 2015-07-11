#coding=utf-8
'''
Created on 2015年7月10日

@author: Anyson
'''

import smtplib

from email.mime.text import MIMEText
from email.header import Header

class Smtp(object):
    u'''smtp服务器的委托类'''
    
    @staticmethod
    def makeEmailMimeText(content = u'', subject = u'', toEmail = u'' , fromEmail = u''):
        msg = MIMEText(content, _subtype='html',_charset='utf-8') 
        msg['Subject'] = subject
        msg['From'] = fromEmail
        msg['to'] = toEmail
        
        return msg.as_string()
        
        
    def __init__(self, userEmail='', password='', smtpServer='', port=25, sslType=0):
        '''
            sslType:0 for no ssl; 1 for SSl; 2 for TLS
        '''
        self.userEmail = userEmail
        self.password = password
        self.smtpServer = smtpServer
        self.port = int(port)
        self.sslType = sslType
            
    def startSmtpServer(self):
        u'''开启smtp服务器'''
        try:
            if self.sslType == 1:
                self.smtp = smtplib.SMTP_SSL(self.smtpServer, int(self.port))
            elif self.sslType == 2:
                self.smtp = smtplib.SMTP(self.smtpServer, int(self.port))
                self.smtp.ehlo()
                self.smtp.starttls()
                self.smtp.ehlo()
            else:
                self.smtp = smtplib.SMTP(self.smtpServer, int(self.port))
            #self.smtp.connect(self.smtpServer, int(self.port))
            self.smtp.set_debuglevel(1)
            self.smtp.login(self.userEmail, self.password)
        #触发连接错误
        except (smtplib.SMTPAuthenticationError,smtplib.SMTPException,Exception):
            raise Exception(u'AuthenticationError')
    
    
    def stopSmtpServer(self):
        u'''停止smtp服务器'''
        if self.smtp:
            self.smtp.quit()
            
    def sendEmail(self, toEmailList=u'', fromEmail=u'', subject=u'', message=u'', ):
        u'''发送邮件，发送前请先设置好发送的内容,也可设置参数发送；先发送参数的message，在发送self.message'''
        try:
            if not hasattr(self,u'smtp'):
                self.startSmtpServer()
               
            if not toEmailList and not self.message:
                return
            if fromEmail == u'':
                fromEmail = self.userEmail
            self.smtp.sendmail(fromEmail, toEmailList , Smtp.makeEmailMimeText(message, subject, toEmailList, fromEmail))
                
        except (smtplib.SMTPAuthenticationError):
            raise Exception(u'AuthenticationError')
        except (smtplib.SMTPServerDisconnected):
            raise Exception(u'ServerDisconnected')
        except (smtplib.SMTPDataError):
            raise Exception(u'DataError')
        except (smtplib.SMTPException):
            raise Exception(u'otherError')#断网
        
        
class TestSmtpServer:
    u'''测试服务器模块'''
    def findSmtpServer(self, smtpServer, port=25, sslType=0):
        u'''查找服务器'''
        print u"开始连接服务器..."
        try:
            
            if sslType == 1:
                self.smtp = smtplib.SMTP_SSL(smtpServer, int(port))
            elif sslType == 2:
                self.smtp = smtplib.SMTP(smtpServer, int(port))
                self.smtp.ehlo()
                self.smtp.starttls()
                self.smtp.ehlo()
            else:
                self.smtp = smtplib.SMTP(smtpServer, int(port))
            self.smtp.set_debuglevel(1)
        except (smtplib.SMTPException, Exception):
            return u'找不到smtp服务器，请确保互联网是否处于连接状态或SMTP服务器设置是否正确。'
        else:
            return u''
    def loginTest(self, userEmail, password):
        u'''连接服务器'''
        print u"开始登陆..."
        try:
            self.smtp.login(userEmail, password)
            self.userEmail = userEmail
            self.password = password
        except (smtplib.SMTPAuthenticationError, Exception):
            return u'连接smtp服务器失败，请确保邮箱和密码是否填写正确。'
        else:
            return u''
        
    def sendTestEmail(self):
        u'''发送测试邮件'''
        try:

            message = MIMEText(u'这是发送工资条小程序自动发送的测试邮件！   ','plain','utf-8')
            message[u'Subject'] = Header(u'测试邮件',u'utf-8')
            message[u'To'] = self.userEmail
            message[u'From'] = self.userEmail
            self.smtp.sendmail(message[u'From'], message[u'To'], message.as_string())
        except (smtplib.SMTPAuthenticationError, smtplib.SMTPSenderRefused, Exception):
            return u'发送测试邮件失败。'
        else:
            return u''
        
if __name__ == "__main__":
    print "Start Test..."
#     testServer = TestSmtpServer()
#     print testServer.findSmtpServer(u"smtp.qq.com", 465, 1)
#     print testServer.loginTest(u"1027509234@qq.com", u"dyh199041@")
#     print testServer.findSmtpServer(u"smtp.office365.com", 587, 2)
#     print testServer.loginTest(u"noreply@sachsen.cc", u"Coxo9070")
#     print testServer.sendTestEmail()


#     smtpInstance = Smtp(u"noreply@sachsen.cc", u"Coxo9070", u"smtp.office365.com", 587, 2)
#     smtpInstance.startSmtpServer()
#     smtpInstance.sendEmail("noreply@sachsen.cc", "noreply@sachsen.cc", "Test", "<b>测试</b>")