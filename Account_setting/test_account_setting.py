from selenium import webdriver
from selenium.webdriver.support.ui import Select
from time import sleep,time
import function_account_setting
from function_account_setting import Register_test,Login_test,Cancellation_account

# 关于注册的测试
driver = webdriver.Chrome()
url = 'https://c.njupt.edu.cn/accounts/register/'
driver.get(url)
start = time()
print ("---------------------------------\n\t\t 开始测试注册功能\n---------------------------------\n")
test_num = 3#指从xlsx文件第三行开始插入数据

function_account_setting.create_xlsx()
input_data = function_account_setting.read_register_xlsx(1)
for i in range(1,int(input_data['row'])):
      driver.find_element_by_css_selector('#top_menu > ul > li:nth-child(1) > a').click()
      input_data = function_account_setting.read_register_xlsx(i+1)
      errors = []
      register_test = Register_test(driver,test_num,input_data)
      register_test.test_register()
      test_num += 1
      sleep(3)
driver.close()
end = time()
test_now = "注册功能的测试时间是:"+str(end-start)+"s"
function_account_setting.add_xlsx(test_now,'A',test_num)
test_num += 1
print("\n注册功能的测试时间是:\n"+str(end-start)+"s"+"\n\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n\t\t 注册功能测试结束"+
      "\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n\n\n\n")


# 关于登录的测试
driver = webdriver.Chrome()
url = 'https://c.njupt.edu.cn/test/accounts/login/'
driver.get(url)
start = time()
print ("---------------------------------\n\t\t 开始测试登录功能\n---------------------------------\n")
input_data = function_account_setting.read_login_xlsx(1)
for i in range(1,int(input_data['row'])):
      driver.find_element_by_css_selector('#top_menu > ul > li:nth-child(2) > a').click()
      input_data = function_account_setting.read_login_xlsx(i+1)
      login_test = Login_test(driver,test_num,input_data)
      login_test.test_login()
      test_num += 1
      sleep(3)
driver.close()
end = time()
test_now = "登录功能的测试时间是:"+str(end-start)+"s"
function_account_setting.add_xlsx(test_now,'A',test_num)
test_num += 1
print("\n登录功能的测试时间是:\n"+str(end-start)+"s"+"\n\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n\t\t 登录功能测试结束"+
      "\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n\n\n\n")

# 关于账户注销的测试
driver = webdriver.Chrome()
url = 'https://c.njupt.edu.cn/test/'
driver.get(url)
start = time()
print ("---------------------------------\n\t\t 开始测试注销功能\n---------------------------------\n")
function_account_setting.add_cookies(driver)
driver.get(url)
cancellation_account = Cancellation_account(driver,test_num)
cancellation_account.test_cancellation_account()
driver.close()
end = time()
test_now = "注销功能的测试时间是:"+str(end-start)+"s"
function_account_setting.add_xlsx(test_now,'A',test_num)
test_num += 1
print("\n注销功能的测试时间是:\n"+str(end-start)+"s"+"\n\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n\t\t 注销功能测试结束"+
      "\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n\n\n\n")

