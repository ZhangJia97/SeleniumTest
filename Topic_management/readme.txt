运行之前请保证电脑已安装python3并且已经安装Selenium （浏览器自动化测试框架）

本程序基于chrome浏览器进行测试，请下载chromedriver.exe并将其放于path环境变量

由于测试时对网速有要求，所以请在网速不慢是进行测试



本程序用于测试程序设计类课程作业平台的题库管理功能

使用前请注意以下几点：

1.在运行前请将更改function_topic_management.py文件中add_cookies函数将其更改成最新的cookies

2.选择题的输入数据请存储于/input/choice_list_data.xlsx中

3.编程题的输入数据请存储于/input/biancheng_list_data.xlsx中

4.填空题的输入数据请存储于/input/tiankong_list_data.xlsx中

5.改错题的输入数据请存储于/input/gaicuo_list_data.xlsx中

6.sample.zip，error.zip分别是正确和错误的测试用例，运行该程序之前请修改function_topic_management.py文件中test_add函数中于文件路径有关的代码，将文件路径改成实际的路径

7.各种类型的题目中，所属课程，知识点一，知识点二请用数字表示，例如所属课程=1指“高级程序设计语言”，
知识点一=1指“高级程序设计语言”中的第一个知识点一“第1章 计算机、C语言与二进制”，
知识点二=1指“第1章 计算机、C语言与二进制”中的第一个二级知识点“1.1 计算机、程序与程序设计语言”

8.输入的数据可以部分不填写，但是请勿整行数据不填写，因为若一行全为空，则程序将会跳过本行

9.最后的实验报告将保存于Topic_management.xlsx文件中，本文件不需要提前新建，程序将自动创建，并在xlsx文件的第一行输入测试时间

10.测试后请将测试报告存储于其他路径，防止下一次测试时将该文件覆盖