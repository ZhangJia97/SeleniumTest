from selenium import webdriver
from selenium.webdriver.support.ui import Select
from time import sleep,time
import function_topic_management
from function_topic_management import Problem_list_test




# 关于题库管理页面的测试
test_num = 3#指从xlsx文件第三行开始插入数据
driver = webdriver.Chrome()
url = 'https://c.njupt.edu.cn/test/'
driver.get(url)
function_topic_management.add_cookies(driver)
driver.get(url)
start = time()
print ("---------------------------------\n\t\t开始测试题库管理页面功能\n---------------------------------\n")

function_topic_management.create_xlsx()

choice_lists = ['choice_list','biancheng_list','tiankong_list','gaicuo_list']
for choice_list in choice_lists:
    # 循环四个题型
    if choice_list == 'choice_list':
        input_data = function_topic_management.read_choice_xlsx(2)#此处的1没有什么意义,主要是为了获取数据的行数，以及为了下一行所需要的input_data变量
    else:
        input_data = function_topic_management.read_problem_xlsx(2,choice_list)

    problem_list_test=Problem_list_test(driver,choice_list,test_num,input_data)
    # 正式测试
    driver.find_element_by_class_name('dcjq-parent-li').click()  # 点击题库管理
    sleep(1)
    driver.find_element_by_id('extend_problem_list').click()
    driver.find_element_by_id(choice_list).click()  # 点击填空题
    problem_list_test.test_select_knowledge()
    test_num += 1
    sleep(1)

    for i in range(1,int(input_data['row'])):
        driver.find_element_by_class_name('dcjq-parent-li').click()  # 点击题库管理
        sleep(1)
        driver.find_element_by_id('extend_problem_list').click()
        driver.find_element_by_id(choice_list).click()  # 点击填空题
        driver.find_element_by_xpath('//*[@id="main-content"]/section/div/div/div/div[1]/a').click()  # 点击添加题目

        if choice_list == 'choice_list':
            input_data = function_topic_management.read_choice_xlsx(i+1)  # 此处的1没有什么意义,主要是为了获取数据的行数，以及为了下一行所需要的input_data变量
        else:
            input_data = function_topic_management.read_problem_xlsx(i+1, choice_list)

        problem_list_test = Problem_list_test(driver, choice_list, test_num, input_data)
        problem_list_test.test_add()
        test_num += 1

    driver.find_element_by_class_name('dcjq-parent-li').click()  # 点击题库管理
    sleep(2)
    driver.find_element_by_id(choice_list).click()  # 返回初始状态
    sleep(1)

    input_data = None
    problem_list_test = Problem_list_test(driver, choice_list,  test_num, input_data)#本步骤主要用于保证test_num增加
    problem_list_test.test_pages()

driver.close()
end = time()
test_now = "题库管理功能的测试时间是:"+str(end-start)+"s"
function_topic_management.add_xlsx(test_now,'A',test_num+1)
print("\n题库管理功能的测试时间是:\n"+str(end-start)+"s"+"\n\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n\t\t题库管理页面功能测试结束"+
      "\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n\n\n\n")
