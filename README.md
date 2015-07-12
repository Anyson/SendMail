# SendMail
发送工资邮件小工具
1.环境相关
	
	a.安装wxPython3.0.2.0（到官网下载安装）：http://wxpython.org/
	b.运行代码前请执行如下命令，安装所用到的库
		pip install -r requirements/requirement.txt
		
2.打包
	进入工程目录执行如下命令
	pyinstaller -F -w -i sendmail/app.ico --distpath=sendmail/client/ --upx-dir=upx/  -p . -n sendEmail sendmail/main.py
	
	完成后进入工程目录下的sendmail/client/即可看到sendEmail.exe。
	copy给他人使用时需要copy client下的所有文件。
		
			