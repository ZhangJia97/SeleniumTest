运行之前请保证电脑已安装python3并且已经安装Selenium （浏览器自动化测试框架）

本程序基于chrome浏览器进行测试，请下载chromedriver.exe并将其放于path环境变量

由于测试时对网速有要求，所以请在网速不慢是进行测试



本程序用于测试程序设计类课程作业平台的账号相关功能

使用前请注意以下几点：

1.在运行前请将更改function_account_setting.py文件中add_cookies函数将其更改成最新的cookies

2.注册的输入数据请存储于/input/register_data.xlsx中

3.登录的输入数据请存储于/input/login_data.xlsx中

4.最后的实验报告将保存于Account_setting.xlsx文件中，本文件不需要提前新建，程序将自动创建，并在xlsx文件的第一行输入测试时间

5.测试后请将测试报告存储于其他路径，防止下一次测试时将该文件覆盖

6.请注意输入的数据的安全，防止测试用的账号密码流出