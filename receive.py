import imaplib
import smtplib
import email
import time
import os
import re
import logging
from email.header import decode_header
from email.utils import parseaddr, formataddr
from email.mime.text import MIMEText

# --------------------- 配置区域 ---------------------
IMAP_SERVER = 'imap.qq.com'          # QQ邮箱IMAP服务器
SMTP_SERVER = 'smtp.qq.com'          # QQ邮箱SMTP服务器
USERNAME = 'xxxxx@qq.com'       # 监控用邮箱
AUTH_CODE = 'xxxxxxx'             # 邮箱授权码（IMAP/SMTP共用）
TARGET_SENDER = 'xxxxxx@163.com'   # 只接受该发件人
CHECK_INTERVAL = 30                  # 检查间隔(秒)
ATTACHMENT_PATH = './attachments'    # 附件保存路径
ALLOWED_COMMANDS = {'1', '2'}       # 允许的命令列表

# --------------------- 日志配置 ---------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_monitor.log'),
        logging.StreamHandler()
    ]
)

# --------------------- 自定义函数 ---------------------
def function_a():
    """命令1对应操作"""
    logging.info("执行函数A：系统状态检查")
    # 在此添加自定义逻辑
    print("Hello from Function A!")

def function_b():
    """命令2对应操作"""
    logging.info("执行函数B：数据备份操作")
    # 在此添加自定义逻辑
    print("Hello from Function B!")

# --------------------- 核心功能 ---------------------
class EmailMonitor:
    def __init__(self):
        self.imap = None
        self.smtp = None
        self._prepare_directories()

    def _prepare_directories(self):
        """创建必要目录"""
        os.makedirs(ATTACHMENT_PATH, exist_ok=True)

    def connect_imap(self):
        """连接IMAP服务器"""
        try:
            self.imap = imaplib.IMAP4_SSL(IMAP_SERVER)
            self.imap.login(USERNAME, AUTH_CODE)
            logging.info("IMAP登录成功")
            return True
        except Exception as e:
            logging.error(f"IMAP连接失败: {str(e)}")
            return False

    def connect_smtp(self):
        """连接SMTP服务器"""
        try:
            self.smtp = smtplib.SMTP_SSL(SMTP_SERVER, 465)
            self.smtp.login(USERNAME, AUTH_CODE)
            logging.info("SMTP登录成功")
            return True
        except Exception as e:
            logging.error(f"SMTP连接失败: {str(e)}")
            return False

    def validate_sender(self, msg):
        """严格验证发件人"""
        sender_header = msg.get("From", "")
        _, sender_email = parseaddr(sender_header)
        return sender_email.lower() == TARGET_SENDER.lower()

    def decode_mail_content(self, content):
        """通用内容解码"""
        try:
            decoded_parts = decode_header(content)
            result = ""
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    result += part.decode(encoding or 'utf-8', 'ignore')
                else:
                    result += str(part)
            return result.strip()
        except:
            return str(content).strip()

    def get_mail_body(self, msg):
        """提取邮件正文"""
        body = ""
        try:
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    disposition = str(part.get("Content-Disposition"))
                    if "attachment" in disposition:
                        continue
                    if content_type == "text/plain":
                        body += self.decode_mail_content(part.get_payload(decode=True))
            else:
                body = self.decode_mail_content(msg.get_payload(decode=True))
        except Exception as e:
            logging.error(f"正文解析失败: {str(e)}")
        return body

    def process_attachments(self, msg):
        """处理邮件附件"""
        try:
            attachments = []
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename = part.get_filename()
                if filename:
                    filename = self.decode_mail_content(filename)
                    filepath = os.path.join(ATTACHMENT_PATH, filename)
                    
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    attachments.append(filepath)
                    logging.info(f"附件保存成功: {filename}")
            return attachments
        except Exception as e:
            logging.error(f"附件处理失败: {str(e)}")
            return []

    def send_confirmation(self, command):
        """发送执行确认邮件"""
        try:
            content = f"""
            您的指令 [{command}] 已成功执行
            执行时间: {time.strftime('%Y-%m-%d %H:%M:%S')}
            """
            
            msg = MIMEText(content, 'plain', 'utf-8')
            msg['From'] = formataddr(("系统监控", USERNAME))
            msg['To'] = TARGET_SENDER
            msg['Subject'] = f"指令执行确认 - {command}"
            
            self.smtp.sendmail(USERNAME, TARGET_SENDER, msg.as_string())
            logging.info(f"确认邮件已发送: {command}")
        except Exception as e:
            logging.error(f"确认邮件发送失败: {str(e)}")

    def process_command(self, content):
        """处理邮件内容中的指令"""
        try:
            # 使用正则精确匹配独立数字
            match = re.search(r'\b([12])\b', content)
            if not match:
                return False
                
            command = match.group(1)
            if command not in ALLOWED_COMMANDS:
                return False

            logging.info(f"检测到有效指令: {command}")
            if command == '1':
                function_a()
            elif command == '2':
                function_b()
            
            if self.connect_smtp():
                self.send_confirmation(command)
            return True
        except Exception as e:
            logging.error(f"指令处理失败: {str(e)}")
            return False

    def process_email(self, msg):
        """处理单封邮件"""
        if not self.validate_sender(msg):
            return

        # 解析邮件内容
        subject = self.decode_mail_content(msg.get("Subject", ""))
        body = self.get_mail_body(msg)
        full_content = f"{subject} {body}"
        
        # 处理附件
        self.process_attachments(msg)
        
        # 处理指令
        if self.process_command(full_content):
            logging.info("邮件处理完成")
        else:
            logging.warning("未发现有效指令")

    def run_monitor(self):
        """主监控循环"""
        while True:
            try:
                if not self.connect_imap():
                    time.sleep(60)
                    continue

                self.imap.select('INBOX')
                status, messages = self.imap.search(
                    None, 
                    f'UNSEEN FROM "{TARGET_SENDER}"'
                )

                if status == 'OK' and messages[0]:
                    for num in messages[0].split():
                        status, data = self.imap.fetch(num, '(RFC822)')
                        if status == 'OK':
                            msg = email.message_from_bytes(data[0][1])
                            self.process_email(msg)
                            self.imap.store(num, '+FLAGS', '\\Seen')

                time.sleep(CHECK_INTERVAL)
            except KeyboardInterrupt:
                logging.info("用户中断监控")
                break
            except Exception as e:
                logging.error(f"监控异常: {str(e)}")
                time.sleep(300)  # 异常等待后重试
            finally:
                if self.imap:
                    try:
                        self.imap.close()
                        self.imap.logout()
                    except:
                        pass

if __name__ == "__main__":
    monitor = EmailMonitor()
    monitor.run_monitor()
