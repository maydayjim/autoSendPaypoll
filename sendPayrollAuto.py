from datetime import datetime
import arrow
import smtplib
from email.mime.multipart import MIMEMultipart    
from email.mime.text import MIMEText    
from email.mime.image import MIMEImage
from email.header import Header
from email.mime.application import MIMEApplication

import xlrd
import xlwt

sender_account = input('请输入发件邮箱：')
pwd = input('请输入发件邮箱的密码：')
smtp = smtplib.SMTP('smtp.qiye.aliyun.com',25) 
smtp.login(sender_account,pwd)

payrollPath= input('请把工资单地址粘贴在这里：')
sheet = xlrd.open_workbook(payrollPath).sheet_by_name('工资')#打开文件
nrows = sheet.nrows #得到行数，从1开始遍历
ncols = sheet.ncols #一共19列，姓名2，邮箱18
rowsNum = int(nrows)
i =1
while i<rowsNum:
	rowsContext = sheet.row_values(i) #获取行内容

	receive_account = rowsContext[18]#获取收件邮箱s

	staffName = rowsContext[2]#获取名字s

	mailAnnex = xlwt.Workbook(encoding = 'ascii')
	wageSheet = mailAnnex.add_sheet('your wage')
	vacationCol=wageSheet.col(6)
	vacationCol.width = 256*20
	row0 = [u'序号',u'部门',u'姓名',u'工作地点',u'正式工资',u'出勤天数',u'休假情况',u'扣除工资',u'应发工资',u'五险一金',u'子女教育',u'继续教育',u'住房贷款',u'住房租金',u'赡养老人',u'应税工资',u'个税',u'实发工资']
	

	for j in range(0,len(row0)):
		wageSheet.write(0,j,row0[j])
		wageSheet.write(1,j,rowsContext[j])


	date = arrow.now()
	dateStr = date.shift(months=-1).format('MM')#获得了上个月的月数
	dateYear = date.format('YYYY')
	annexName = staffName+dateYear+'年'+dateStr+'月薪资明细_'+'.xls'
	mailAnnex.save(annexName) #这里可以打包好附件信息了

	message = MIMEMultipart()
	message['From'] = Header('Soul C&B', 'utf-8')
	message['To'] =  Header(staffName, 'utf-8')
	message['Subject'] = Header(dateYear+'年'+dateStr+'月'+'薪资明细', 'utf-8')
	message.attach(MIMEText('Dear，请查收附件'+dateYear+'年'+dateStr +'月薪资明细。\n\n如有疑问请及时联系人力资源部XXX：\nxxx@XXX.cn', 'plain', 'utf-8'))

	xlsFile = annexName
	xlsApart = MIMEApplication(open(xlsFile, 'rb').read())
	xlsApart.add_header('Content-Disposition', 'attachment', filename=xlsFile)

	message.attach(xlsApart)#把附加加到邮件中了

	smtp.sendmail(sender_account, receive_account, message.as_string())#发送邮件

	print('已将薪资明细发送给 '+staffName)

	i+=1

print('********')
print('发送完毕！')

smtp.quit()
