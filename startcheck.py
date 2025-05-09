import cv2
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# import smtplib #发送邮件
import smtplib
from smtplib import SMTP
import time
import psutil
import subprocess
import sys
import pyautogui

#邮箱配置
host = 'smtp.qq.com' #邮箱的接口
port = '465' #端口
pwd = 'xxxxxxxxxxx' #授权码
sender = 'xxxxxxx@qq.com' #发送方
receiver = "xxxxxx@163.com" #接收方

#时间配置
current_time = time.strftime("%Y-%m-%d-%H-%M-%S")  #显示发送邮件时间

# 检测笔记本电源状态
def check_power_status():
    battery = psutil.sensors_battery()
    if battery is None:
        return "未检测到电池（可能是台式机）", "N/A"
    
    # 检测电源状态
    plugged = "已接通电源" if battery.power_plugged else "未接通电源"
    # 获取电量百分比
    percent = f"{battery.percent}%"
    return plugged, percent

#检测程序是否在运行状态
def check_desk():
    def check_multiple_processes(target_names):
        running_processes = set()
        for proc in psutil.process_iter(['name']):
            current_name = proc.info['name']
            if current_name and current_name.lower() in target_names:
                running_processes.add(current_name.lower())
        return running_processes

    # 目标进程名列表（需根据实际进程名调整）
    targets = {'todesk.exe', 'rustdesk.exe'}  
    result = check_multiple_processes(targets)

    lines = []  # 初始化空列表

    for name in targets:
        status = "是" if name in result else "否"
        lines.append(f"{name} {status}")  # 将每行内容加入列表
    result_checkrun = "\n".join(lines)  # 合并列表为字符串，每行用换行符分隔
    return result_checkrun
result_run = check_desk() #获取应用运行状态

#设置邮件格式
def SetMsg(screenshot_path):
    msg = MIMEMultipart('mixed')
    #标题
    msg['Subject'] = current_time + '电脑正在运行'
    msg['From'] = sender
    msg['To'] = receiver
    #邮件正文内容
    text = '电源状态：'+ power_status + "\n" + '电量' + power_percent + "\n" + '运行中的进程:' + "\n" + result_run + "\n"
    text_plain = MIMEText(text,'plain','utf-8') #正文转码
    msg.attach(text_plain)

    # 添加图片附件
    with open(screenshot_path, 'rb') as f:
        img_data = f.read()
    image = MIMEImage(img_data, name=screenshot_path)
    image.add_header('Content-Disposition', 'attachment', filename=screenshot_path)
    msg.attach(image)

    return msg.as_string()

#发送邮件
def SendEmail(msg):
    try:
        smtp = smtplib.SMTP_SSL(host,port) #创建一个邮件服务
        # smtp.connect(host)
        smtp.login(sender,pwd)
        smtp.sendmail(sender,receiver,msg)
        time.sleep(3)
        smtp.quit() #退出邮件服务
    except smtplib.SMTPException as e:
        print("e")

#设置每隔三十分钟发送一次正在运行邮件
def mail():
        #设置邮件格式
    msg = SetMsg(screenshot_path)
        #发送邮件
    SendEmail(msg)
    num=1
    current_func = sys._getframe().f_code.co_name
    while True:
        print(num)
        print(current_time+"  #  "+current_func)
        time.sleep(1800)  # 30分钟
            #设置邮件格式
        msg = SetMsg(screenshot_path)
            #发送邮件
        SendEmail(msg)
        num +=1

#执行应用运行命令
def run_desk():
    todesk= [r"C:\Program Files\Google\Chrome\Application\chrome.exe"] #配置应用地址
    rustdesk= [r"D:\1"]
    subprocess.run(rustdesk)

def get_pic():
    current_time = time.strftime("%Y-%m-%d-%H-%M-%S")  # 每次生成当前时间
    screenshot_path = f"{current_time}_screenshot.png"
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_path)
    return screenshot_path

if __name__ == '__main__':
    power_status, power_percent = check_power_status()  # 获取电源电池状态
    check_desk()
    screenshot_path = get_pic()
    mail()




