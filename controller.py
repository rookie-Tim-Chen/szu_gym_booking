# -*- coding: utf-8 -*-
import time
import imaplib
import ssl
from datetime import datetime
import sys
import io
import email
import re
from email.header import decode_header
from collections import deque
from email.utils import parsedate_to_datetime

# 控制台编码设置
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

CONFIG = {
    "email": "443799744@qq.com",
    "auth_code": "pfgcfubiucwsbggb",
    "imap_server": "imap.qq.com",
    "port": 993,
    "command_format": r"订场-(\d+)-(\d{1,2})-(\d{1,2})",
    "max_retry": 3,
    "poll_interval": 10
}

execution_log = deque(maxlen=100)


def log_message(message, status=None):
    """日志函数（保持不变）"""
    status_map = {
        "success": ("[OK] ", "\033[92m"),
        "error": ("[ERROR] ", "\033[91m"),
        "warning": ("[WARNING] ", "\033[93m"),
        None: ("", "")
    }
    symbol, color = status_map.get(status, ("", ""))
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{color}[{timestamp}] {symbol}{message}\033[0m", flush=True)
    with open("imap_auto.log", "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {symbol}{message}\n")


def decode_header_field(header):
    """邮件头解码（保持不变）"""
    try:
        return ''.join(
            p.decode(c or 'utf-8', errors='replace') if isinstance(p, bytes) else str(p)
            for p, c in decode_header(header or "")
        )
    except Exception as e:
        log_message(f"头信息解码失败: {str(e)}", "error")
        return str(header)


def validate_time_range(start, end):
    """时间验证（保持不变）"""
    try:
        s, e = int(start), int(end)
        return 0 <= s <= 24 and 0 <= e <= 24 and s < e
    except ValueError:
        return False


def parse_booking_command(subject):
    """命令解析（保持不变）"""
    match = re.search(CONFIG["command_format"], subject)
    if not match:
        return None

    day = int(match.group(1))
    start_time = f"{int(match.group(2)):02d}"
    end_time = f"{int(match.group(3)):02d}"

    if not (1 <= day <= 7):
        log_message(f"无效天数参数: {day}", "warning")
        return None
    if not validate_time_range(start_time, end_time):
        log_message(f"无效时间范围: {start_time}-{end_time}", "warning")
        return None

    return {
        "choose_day": day,
        "choose_time": f"{start_time}-{end_time}",
        "command_hash": hash(f"{day}{start_time}{end_time}")
    }


def execute_booking(command):
    """执行预定（保持不变）"""
    log_message(f"执行预定：第{command['choose_day']}天 {command['choose_time']}", "success")
    return True


def process_unseen_emails():
    """处理未读邮件（关键修改部分）"""
    context = ssl.create_default_context()
    try:
        with imaplib.IMAP4_SSL(CONFIG["imap_server"], CONFIG["port"], ssl_context=context) as mail:
            mail.login(CONFIG["email"], CONFIG["auth_code"])
            mail.select("INBOX")

            # 获取未读邮件UID列表
            status, uids = mail.uid('search', None, 'UNSEEN')
            if status != 'OK' or not uids[0]:
                return

            valid_emails = []
            current_time = datetime.now()

            # 第一阶段：筛选有效邮件
            for uid in uids[0].split():
                status, data = mail.uid('fetch', uid, '(BODY.PEEK[HEADER.FIELDS (DATE)])')
                if status != 'OK' or not data[0]:
                    continue

                try:
                    msg = email.message_from_bytes(data[0][1])
                    date_str = msg.get('Date', '')

                    # 解析邮件时间（强制转换为本地时间）
                    mail_date = parsedate_to_datetime(date_str)
                    mail_date = mail_date.astimezone().replace(tzinfo=None)

                    # 有效性校验
                    time_diff = (current_time - mail_date).total_seconds()

                    # 排除未来邮件和过期邮件
                    if time_diff < -60:  # 邮件时间比当前时间晚60秒以上
                        log_message(f"忽略未来邮件: {mail_date} (当前{current_time})", "warning")
                        mail.uid('store', uid, '+FLAGS', '\\Seen')  # 标记未来邮件为已读
                        continue

                    if 0 <= time_diff <= 60:
                        valid_emails.append((uid, mail_date))

                except Exception as e:
                    log_message(f"日期解析异常: {str(e)} UID:{uid.decode()}", "warning")
                    continue

            # 第二阶段：处理最新有效邮件
            if valid_emails:
                # 按时间排序找最新
                valid_emails.sort(key=lambda x: x[1], reverse=True)
                latest_uid, latest_time = valid_emails[0]

                # 获取完整邮件内容
                status, data = mail.uid('fetch', latest_uid, '(RFC822)')
                if status != 'OK':
                    return

                # 解析和执行逻辑
                msg = email.message_from_bytes(data[0][1])
                command = parse_booking_command(decode_header_field(msg['Subject']))

                if command and command['command_hash'] not in execution_log:
                    if execute_booking(command):
                        execution_log.append(command['command_hash'])
                        mail.uid('store', latest_uid, '+FLAGS', '\\Seen')
                        log_message(f"成功处理并标记邮件: UID {latest_uid.decode()}", "success")

                        # 标记其他有效邮件为已读（防止重复处理）
                        for uid, _ in valid_emails[1:]:
                            mail.uid('store', uid, '+FLAGS', '\\Seen')
                            log_message(f"标记关联邮件: UID {uid.decode()}", "warning")

            # 强制刷新邮箱状态
            mail.close()
            mail.select("INBOX")

    except Exception as e:
        log_message(f"邮件处理异常: {str(e)}", "error")


if __name__ == "__main__":
    retry_count = 0
    while retry_count < CONFIG["max_retry"]:
        try:
            process_unseen_emails()
            time.sleep(CONFIG["poll_interval"])
            retry_count = 0
        except KeyboardInterrupt:
            log_message("用户主动终止程序", "warning")
            break
        except Exception as e:
            retry_count += 1
            log_message(f"运行时异常 ({retry_count}/{CONFIG['max_retry']}): {str(e)}", "error")
            time.sleep(10)



