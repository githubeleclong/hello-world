from demo_deautifulsoup import GetAppointmentTime_Shenzhen
from webrequesttest import GetAppointmentTime_Shanghai
import pandas as pd

# 深圳市场
pd_shenzhenmainnew = GetAppointmentTime_Shenzhen()
# print(pd_shenzhenmainnew)

pd_shenzhenmainold = pd.read_excel(r'C:\\深圳市场预约时间表.xls')
# print(pd_shenzhenmainold)
pd_shenzhenmainnew.to_excel(r'C:\\深圳市场预约时间表.xls',sheet_name='深圳主板')
pd_shenzhenmainnew=pd.read_excel(r'C:\\深圳市场预约时间表.xls')
pd_shenzhenmainold=pd_shenzhenmainold.append(pd_shenzhenmainnew)

result_shenzhen = pd_shenzhenmainold.drop_duplicates(keep=False)
# print(result)
# print(len(result))
result_shenzhen = result_shenzhen.drop_duplicates(2,keep='last')
# print(result)
dfsize_shenzhen=len(result_shenzhen)
msg=''
for i in range(0,dfsize_shenzhen):
    msg += str(result_shenzhen.iloc[i][1])
    msg += ':'
    msg += result_shenzhen.iloc[i][2]
    msg += '\n'
    msg += '首次预约时间：'
    msg += result_shenzhen.iloc[i][3]
    msg += '\n'
    msg += '第一次变更时间：'
    msg += result_shenzhen.iloc[i][4]
    msg +='\n'
    msg += '第二次变更时间：'
    msg += result_shenzhen.iloc[i][5]
    msg += '\n'
    msg += '第三次变更时间：'
    msg += result_shenzhen.iloc[i][6]
    msg += '\n'
    msg += '\n'

# 上海市场
pd_shanghaimainnew1=GetAppointmentTime_Shanghai(count=10000)
pd_shanghaimainold=pd.read_excel(r'C:\\上海市场预约时间表.xls')
pd_shanghaimainnew1.to_excel(r'C:\\上海市场预约时间表.xls',sheet_name='上海主板')
pd_shanghaimainnew1=pd.read_excel(r'C:\\上海市场预约时间表.xls')
# print(pd_shanghaimainnew)
# pd_shanghaimainold=pd_shanghaimainold.where(pd_shanghaimainold.notnull(), 'blank')
# print(pd_shanghaimainold)
pd_shanghaimainold = pd_shanghaimainold.append(pd_shanghaimainnew1)
#pd_shanghaimainold.to_excel(r'C:\\testappendrst上海市场预约时间表.xls',sheet_name='上海主板')

result_shanghai = pd_shanghaimainold.drop_duplicates(keep=False)
result_shanghai = result_shanghai.drop_duplicates('companyCode',keep='last')
dfsize_shanghai=len(result_shanghai)
num=0
for i in range(0,dfsize_shanghai):
    msg += str(result_shanghai.iloc[i]['companyCode'])
    msg += ':'
    msg += result_shanghai.iloc[i]['companyAbbr']
    msg += '\n'
    msg += '首次预约时间：'
    msg += result_shanghai.iloc[i]['publishDate0']
    msg += '\n'
    msg += '第一次变更时间：'
    msg += result_shanghai.iloc[i]['publishDate1']
    msg +='\n'
    msg += '第二次变更时间：'
    msg += result_shanghai.iloc[i]['publishDate2']
    msg += '\n'
    msg += '第三次变更时间：'
    msg += result_shanghai.iloc[i]['publishDate3']
    msg += '\n'
    msg += '\n'
    num+=1

import sendemail
if msg == '':
    msg ='年报预约时间没有变更'
print(msg)
sendemail.send_email(msg)