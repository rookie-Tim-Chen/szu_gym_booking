from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from datetime import datetime
import logging
import traceback
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from threading import Event
from multiprocessing import Process
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--choose_day', type=int, required=True)
parser.add_argument('--choose_time', type=str, required=True)
args = parser.parse_args()

# 配置日志
logging.basicConfig(
    level=logging.INFO,  # 可根据需要调整日志级别
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 配置参数
CONFIG = {
    "login_url": "https://ehall.szu.edu.cn/qljfwapp/sys/lwSzuCgyy/index.do#/sportVenue",
    "user_name": "",
    "password": "",
    "choose_day": args.choose_day,  # 当天为1，第二天为2
    "choose_time": args.choose_time,  # 24小时制，例如晚上8-9点输入 "20-21"
    "date_time": "12:29:30",  # 定时执行时间
    "try_other_times": True,  # 当前时间段不可用时尝试其他时间段
    "time_range_offset": 1,  # 尝试的前后时间段数量
    "wait_time": 2,  # 选择时间段的等待时间，单位秒
    "max_attempts": 2000,  # 最大尝试次数
    "email": {                              #开启SMTP 参考：https://laowangblog.com/qq-mail-smtp-service.html
        "from_email": "",  # 替换为您的 QQ 邮箱        
        "from_password": "",  # 替换为您的 QQ 邮箱授权码
        "to_email": ""       #接受信息的邮箱 ，学校的邮箱也行
    },
    "num_processes": 3  # 启动的并行实例数，根据需要调整
}
        #账户中有余额才能付款
def wait_until(target_time):
    event = Event()
    while True:
        now = datetime.now().time()
        if now >= target_time:
            break
        remaining = (datetime.combine(datetime.today(), target_time) - datetime.now()).total_seconds()
        event.wait(timeout=min(remaining, 0.5))

def send_email(subject, body, to_email):
    from_email = CONFIG["email"]["from_email"]
    from_password = CONFIG["email"]["from_password"]

    # 创建邮件对象
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # 添加邮件正文
    msg.attach(MIMEText(body, 'plain'))

    try:
        # 连接到 QQ 邮箱服务器
        server = smtplib.SMTP_SSL('smtp.qq.com', 465)  # 使用 SSL 连接
        server.set_debuglevel(0)  # 关闭调试输出
        server.login(from_email, from_password)

        # 发送邮件
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        logger.info("邮件发送成功")
    except Exception as e:
        logger.error(f"邮件发送失败: {e}")

def convert_time_range_to_number(time_range):
    try:
        start_hour = int(time_range.split('-')[0])
        return start_hour - 7 if 8 <= start_hour <= 21 else None
    except (ValueError, IndexError) as e:
        logger.error(f"时间范围转换失败: {e}")
        return None

def select_element(driver, by, selector, timeout=10, description="元素"):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((by, selector))
        )
        return element
    except TimeoutException:
        logger.warning(f"{description} 在 {timeout} 秒内不可点击")
        return None

def click_element(driver, by, selector, timeout=20, description="元素"):
    element = select_element(driver, by, selector, timeout, description)
    if element:
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        driver.execute_script("arguments[0].click();", element)
        logger.info(f"成功点击 {description} 按钮")
        return True
    return False

def set_input_value(driver, by, selector, value, timeout=10, description="输入框"):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((by, selector))
        )
        element.clear()
        element.send_keys(value)
        logger.info(f"成功设置 {description} 的值")
        return True
    except TimeoutException:
        logger.error(f"{description} 在 {timeout} 秒内不可见")
        return False

def switch_to_iframe(driver, iframe_selector, timeout=10):
    try:
        iframe = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, iframe_selector))
        )
        driver.switch_to.frame(iframe)
        logger.info(f"成功切换到 iframe: {iframe_selector}")
        return True
    except TimeoutException:
        logger.error(f"在 {timeout} 秒内未找到 iframe: {iframe_selector}")
        return False

def click_reservation_buttons(driver):
    """
    点击“场馆预约按钮”和“羽毛球馆”按钮的函数。
    """
    # 点击场馆预约按钮
    if not click_element(driver, By.CSS_SELECTOR, "#sportVenue > div:nth-child(1) > div > div:nth-child(1)", description="场馆预约按钮"):
        raise Exception("场馆预约按钮点击失败")

    # 点击羽毛球馆
    if not click_element(driver, By.CSS_SELECTOR,
                         "#sportVenue > div:nth-child(2) > div.group-11 > div.overlap-6 > div:nth-child(1) > div",
                         description="羽毛球馆"):
        raise Exception("羽毛球馆按钮点击失败")

def select_venue_and_date(driver):
    try:
        # 点击场馆预约按钮和羽毛球馆
        click_reservation_buttons(driver)

        # 选择日期
        for attempt in range(CONFIG["max_attempts"]):
            try:
                date_selector = f"#apply > div.rectangle-2 > div:nth-child(4) > div:nth-child({CONFIG['choose_day']}) > label"
                date_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, date_selector))
                )
                date_button.click()
                logger.info("日期选择成功")
                break
            except StaleElementReferenceException:
                logger.warning("元素变得陈旧，重新尝试定位元素")
                continue  # 重新尝试定位元素
            except TimeoutException:
                logger.warning(f"尝试第 {attempt + 1} 次，日期按钮不可点击，正在刷新页面")
                try:
                    driver.refresh()  # 刷新当前页面
                    logger.info("页面已刷新")
                    time.sleep(2)  # 等待页面刷新完成
                    # 重新点击场馆预约按钮和羽毛球馆
                    click_reservation_buttons(driver)
                except Exception as refresh_e:
                    logger.error(f"刷新页面失败或重新点击按钮失败: {refresh_e}")
                if attempt + 1 >= CONFIG["max_attempts"]:
                    logger.error("超过最大尝试次数，无法找到日期按钮")
                    raise Exception("日期按钮选择失败")
                # 继续尝试下一次循环
    except Exception as e:
        logger.error(f"选择场馆和日期过程中出现问题: {e}")
        logger.error(traceback.format_exc())
        raise  # 如果选择场馆或日期失败，终止脚本

def select_available_time_slot(driver, initial_time, try_other_times, time_range_offset, max_attempts, wait_time):
    if initial_time is None:
        logger.error("无效的初始时间")
        return False

    for attempt in range(1, max_attempts + 1):
        logger.info(f"尝试第 {attempt} 次选择时间段")
        # 生成要尝试的时间段列表
        time_ranges = [initial_time]
        if try_other_times:
            time_ranges += list(range(max(initial_time - time_range_offset, 1), min(initial_time + time_range_offset + 1, 15)))
            time_ranges = list(dict.fromkeys(time_ranges))  # 去重并保持顺序

        available_slot_found = False

        for time_slot in time_ranges:
            try:
                selector = f"#apply > div.rectangle-2 > div:nth-child(6) > div:nth-child({time_slot}) > label"
                time_button = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                button_text = time_button.text
                logger.debug(f"检查时间段 {button_text}")
                if "(可预约)" in button_text:
                    driver.execute_script("arguments[0].scrollIntoView(true);", time_button)
                    driver.execute_script("arguments[0].click();", time_button)
                    logger.info(f"成功选择时间段：{button_text}")
                    available_slot_found = True
                    break  # 退出时间段循环
                else:
                    logger.debug(f"时间段 {button_text} 不可预约，尝试下一个")
            except TimeoutException:
                logger.debug(f"无法找到时间段按钮 {time_slot}")

        if available_slot_found:
            return True  # 成功找到并选择了可预约的时间段

        if attempt < max_attempts:
            logger.info(f"未找到可预约的时间段，等待 {wait_time} 秒后重试...")
            try:
                # 使用 JavaScript 刷新页面并重新选择场馆和日期
                driver.execute_script("location.reload();")  # 如果无法部分刷新，可以保留全刷新
                select_venue_and_date(driver)  # 重新选择场馆和日期
            except Exception as e:
                logger.error(f"刷新页面后重新选择场馆和日期失败: {e}")
                return False
            time.sleep(wait_time)
        else:
            logger.error("在指定范围内没有找到可预约的时间段，已达到最大尝试次数")

    return False

def book_venue(driver, time_number):
    booked_venue = None  # 初始化已预订的场地变量
    try:
        select_venue_and_date(driver)

        # 选择时间段
        if not select_available_time_slot(
            driver,
            initial_time=time_number,
            try_other_times=CONFIG["try_other_times"],
            time_range_offset=CONFIG["time_range_offset"],
            max_attempts=CONFIG["max_attempts"],
            wait_time=CONFIG["wait_time"]
        ):
            raise Exception("预约失败，指定范围内所有时间段都不可用")

        # 选择场地
        for i in range(1, 32):
            try:
                button_selector = f"#apply > div.rectangle-2 > div:nth-child(10) > div:nth-child({i})"
                place_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, button_selector))
                )
                # 获取按钮文本
                button_text = place_button.text
                if "可预约" in button_text and "已满员" not in button_text:
                    place_button.click()
                    logger.info(f"成功点击场地 {i} 按钮，文本：{button_text}")
                    # 提取场地号，例如“羽毛球场A7”中的“A7”
                    if "羽毛球场" in button_text:
                        booked_venue = button_text.replace("羽毛球场", "").strip().split("(")[0]
                    else:
                        booked_venue = button_text.strip().split("(")[0]
                    break
                else:
                    logger.debug(f"场地 {i} 按钮不可点击，文本：{button_text}")
            except TimeoutException:
                logger.debug(f"场地 {i} 不可点击，跳过")

        if not booked_venue:
            logger.error("未能成功选择场地")
            raise Exception("选择场地失败")

        # 提交预订
        if not click_element(driver, By.CSS_SELECTOR, "#apply > div.rectangle-2 > div:nth-child(13) > button:nth-child(2)", description="提交预订"):
            raise Exception("提交预订按钮点击失败")

        logger.info("预订提交成功")

        # 点击未支付
        if not click_element(driver, By.CSS_SELECTOR, "#row0myBookingInfosTable > td.jqx-cell.jqx-grid-cell.jqx-item.jqx-center-align > a.j-row-pay", description="未支付"):
            raise Exception("点击“未支付”失败")
        logger.info("点击“未支付”成功")

        # 点击(剩余金额)支付
        if not click_element(driver, By.CSS_SELECTOR, "#buttons > button", description="剩余金额支付"):
            raise Exception("点击“支付”失败")
        logger.info("点击“支付”成功")

        # 发送邮件通知，包含已预订的场地号
        if booked_venue:
            email_body = f"您的场地预订已成功完成。\n\n预订详情如下：\n场地号：{booked_venue}\n时间段：{CONFIG['choose_time']}\n第:{CONFIG['choose_day']}天\n预定账号：{CONFIG['user_name']}"
        else:
            email_body = "您的场地预订已成功完成。"

        send_email(
            subject="预订成功通知",
            body=email_body,
            to_email=CONFIG["email"]["to_email"]
        )

    except Exception as e:
        logger.error(f"预订过程中出现问题: {e}")
        logger.error(traceback.format_exc())
        # 根据需要，可以选择是否重新尝试
        raise  # 终止当前实例

def main():
    # 初始化参数
    time_number = convert_time_range_to_number(CONFIG["choose_time"])
    if time_number is None:
        logger.error("无效的时间范围，程序终止")
        return

    # 配置 Chrome 选项
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # 保持无头模式以提高速度
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--disable-images')
    prefs = {"profile.managed_default_content_settings.images": 2}
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)

    try:
        # 等待直到特定时间
        target_time = datetime.strptime(CONFIG["date_time"], '%H:%M:%S').time()
        logger.info(f"等待直到目标时间 {target_time.strftime('%H:%M:%S')}")
        wait_until(target_time)
        logger.info("到达目标时间，开始执行程序")

        # 打开登录页面
        driver.get(CONFIG["login_url"])
        logger.info("打开登录页面")

        # 登录部分
        try:
            # 设置用户名
            if not set_input_value(driver, By.CSS_SELECTOR, "#username", CONFIG["user_name"], description="用户名"):
                raise Exception("用户名输入失败")

            # 使用 JavaScript 设置密码，避免安全措施阻止 send_keys
            try:
                password_script = "document.querySelector('#password').value = arguments[0];"
                driver.execute_script(password_script, CONFIG["password"])
                logger.info("成功设置密码")
            except Exception as e:
                logger.error(f"设置密码失败: {e}")
                raise

            # 点击登录按钮
            if not click_element(driver, By.CSS_SELECTOR, "#login_submit", description="登录"):
                raise Exception("登录按钮点击失败")

            # 等待登录完成，确保进入主页面
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#sportVenue > div:nth-child(1) > div > div:nth-child(1)"))
                )
                logger.info("登录成功，进入主页面")
            except TimeoutException:
                logger.error("登录后未能检测到主页面元素，登录可能失败或页面加载超时")
                raise
        except Exception as e:
            logger.error(f"登录过程中出现问题: {e}")
            logger.error(traceback.format_exc())
            raise  # 如果登录失败，终止脚本

        # 预订场地
        try:
            book_venue(driver, time_number)
        except Exception as e:
            logger.error(f"预订过程中出现严重问题: {e}")
            logger.error(traceback.format_exc())
            # 根据需要，可以选择是否继续或终止脚本

    except Exception as main_e:
        logger.error(f"脚本执行过程中出现严重问题: {main_e}")
        logger.error(traceback.format_exc())

    finally:
        logger.info("等待1分钟后关闭浏览器")
        time.sleep(60)  # 延迟1分钟，注意修正为60秒
        driver.quit()
        logger.info("关闭浏览器")

def run_booking_instance():
    try:
        main()
    except Exception as e:
        logger.error(f"预订实例遇到错误: {e}")
        logger.error(traceback.format_exc())

if __name__ == '__main__':
    # 启动多个并行实例以增加成功率
    processes = []
    for _ in range(CONFIG["num_processes"]):
        p = Process(target=run_booking_instance)
        p.start()
        processes.append(p)

    for p in processes:
        p.join()

