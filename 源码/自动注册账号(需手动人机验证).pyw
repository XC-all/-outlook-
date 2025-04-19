import random
import string
import time
import sys
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, WebDriverException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import names  # For generating random names
import threading

def generate_random_password(length=12):
    """Generate a random password with letters, numbers, and special characters."""
    # 确保密码包含至少一个大写字母、一个小写字母、一个数字和一个特殊字符
    lowercase = string.ascii_lowercase
    uppercase = string.ascii_uppercase
    digits = string.digits
    special = "!@#$%^&*"
    
    # 随机选择各种字符
    pwd = [
        random.choice(lowercase),
        random.choice(uppercase),
        random.choice(digits),
        random.choice(special)
    ]
    
    # 填充剩余长度
    pwd.extend(random.choice(lowercase + uppercase + digits + special) for _ in range(length - 4))
    
    # 打乱顺序
    random.shuffle(pwd)
    
    return ''.join(pwd)

def generate_random_email(length=8):
    """Generate a random email username."""
    chars = string.ascii_lowercase + string.digits
    username = ''.join(random.choice(chars) for _ in range(length))
    return username

def generate_random_birth_date():
    """Generate a random birth date for someone between 18 and 60 years old."""
    year = random.randint(1964, 2004)
    month = random.randint(1, 12)
    day = random.randint(1, 28)  # Using 28 to avoid month/day validation issues
    return {'year': year, 'month': month, 'day': day}

def wait_and_click(driver, locator, timeout=30, sleep_after=1, description="元素"):
    """等待元素可点击，然后点击它，处理常见的异常情况"""
    try:
        wait = WebDriverWait(driver, timeout)
        element = wait.until(EC.element_to_be_clickable(locator))
        print(f"找到{description}: {element.text if element.text else '无文本'}")
        
        # 使用JavaScript点击，更可靠
        driver.execute_script("arguments[0].click();", element)
        print(f"成功点击{description}")
        
        if sleep_after > 0:
            time.sleep(sleep_after)
        return True
    except Exception as e:
        print(f"点击{description}时出错: {e}")
        return False

def try_multiple_locators(driver, locators, action_type="click", text=None, timeout=2, sleep_after=0.1, description="元素"):
    """尝试多个定位器，直到找到并执行操作"""
    for locator in locators:
        try:
            if action_type == "click":
                try:
                    wait = WebDriverWait(driver, timeout)
                    element = wait.until(EC.element_to_be_clickable(locator))
                    driver.execute_script("arguments[0].click();", element)
                    print(f"成功点击{description}")
                    time.sleep(sleep_after)
                    return True
                except:
                    continue
            elif action_type == "input" and text:
                try:
                    wait = WebDriverWait(driver, timeout)
                    element = wait.until(EC.visibility_of_element_located(locator))
                    element.clear()
                    element.send_keys(text)
                    print(f"已在{description}中输入: {text}")
                    time.sleep(sleep_after)
                    return True
                except:
                    continue
        except:
            continue
    
    print(f"所有尝试的定位器都失败了，无法在页面上{action_type} {description}")
    return False

def register_outlook_account(auto_mode=True):
    """
    Outlook邮箱自动注册功能
    
    参数:
        auto_mode: 设置为True时尝试全自动注册，False时使用手动辅助模式
    """
    # Setup Chrome in incognito mode
    chrome_options = Options()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--start-maximized")
    
    # Optional: Add these arguments to make the browser more stable
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    # 创建账户信息
    first_name = names.get_first_name()
    last_name = names.get_last_name()
    email_username = generate_random_email(8)
    password = generate_random_password(12)
    birth_date = generate_random_birth_date()
    
    # 创建完整邮箱地址（仅用于显示和保存，而非输入）
    domains = ["outlook.com", "hotmail.com"]
    email = f"{email_username}@{random.choice(domains)}"
    
    # 打印账号信息
    print("\n============= 账号信息 =============")
    print(f"邮箱: {email}")
    print(f"密码: {password}")
    print(f"姓名: {first_name} {last_name}")
    print(f"生日: {birth_date['year']}-{birth_date['month']}-{birth_date['day']}")
    print("==================================\n")
    
    # 保存账号信息到文件
    with open("registered_accounts.txt", "a", encoding="utf-8") as f:
        f.write(f"邮箱: {email}\n")
        f.write(f"密码: {password}\n")
        f.write(f"姓名: {first_name} {last_name}\n")
        f.write(f"生日: {birth_date['year']}-{birth_date['month']}-{birth_date['day']}\n")
        f.write("-" * 40 + "\n")
    print("账号信息已保存至 registered_accounts.txt")
    
    # 启动浏览器
    driver = None
    try:
        print("正在初始化浏览器...")
        
        # 尝试多种方式启动浏览器
        success = False
        
        # 方法1: 使用ChromeDriverManager
        if not success:
            try:
                print("尝试方法1: 使用ChromeDriverManager...")
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)
                print("浏览器初始化成功！")
                success = True
            except Exception as e:
                print(f"方法1失败: {e}")
        
        # 方法2: 直接初始化Chrome
        if not success:
            try:
                print("尝试方法2: 直接初始化Chrome...")
                driver = webdriver.Chrome(options=chrome_options)
                print("浏览器直接初始化成功！")
                success = True
            except Exception as e:
                print(f"方法2失败: {e}")
        
        # 方法3: 使用Service但不使用ChromeDriverManager
        if not success:
            try:
                print("尝试方法3: 使用Service但不使用自动下载...")
                service = Service()  # 使用默认路径
                driver = webdriver.Chrome(service=service, options=chrome_options)
                print("浏览器使用默认Service初始化成功！")
                success = True
            except Exception as e:
                print(f"方法3失败: {e}")
        
        # 方法4: 提示用户指定ChromeDriver路径
        if not success:
            print("\n无法自动初始化浏览器。")
            print("请确保已安装Chrome浏览器，并下载对应版本的ChromeDriver。")
            print("您可以从以下网址下载ChromeDriver: https://sites.google.com/chromium.org/driver/")
            
            use_custom_path = input("是否要手动指定ChromeDriver路径? (y/n): ")
            if use_custom_path.lower() == 'y':
                driver_path = input("请输入ChromeDriver的完整路径: ")
                try:
                    service = Service(executable_path=driver_path)
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    print("使用自定义路径初始化浏览器成功！")
                    success = True
                except Exception as e:
                    print(f"使用自定义路径失败: {e}")
        
        # 如果所有方法都失败
        if not success:
            print("\n所有浏览器初始化方法都失败。请检查：")
            print("1. Chrome浏览器是否正确安装")
            print("2. ChromeDriver是否与您的Chrome版本兼容")
            print("3. ChromeDriver是否在系统PATH中，或者可以手动指定路径")
            input("按Enter键退出...")
            return
        
        actions = ActionChains(driver)
        
        # 访问注册页面
        print("正在访问注册页面...")
        url = "https://signup.live.com/signup?lcid=1033&wa=wsignin1.0&rpsnv=174&ct=1743765553&rver=7.5.2211.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26signup%3d1%26cobrandid%3dab0455a0-8d03-46b9-b18b-df2f57b9e44c%26RpsCsrfState%3d6a7bc9db-94fa-01bf-266f-6ac2d3cb08c9&id=292841&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=ab0455a0-8d03-46b9-b18b-df2f57b9e44c&lic=1&uaid=7eb675822a5449da9127d18443476f5d"
        driver.get(url)
        print("页面加载中...")
        
        # 等待页面加载
        time.sleep(5)
        
        # ===== 步骤1: 点击"同意并继续"按钮 =====
        print("\n步骤1: 处理个人数据导出许可对话框")
        
        # 尝试多种方式定位"同意并继续"按钮
        consent_button_locators = [
            (By.XPATH, "//button[contains(text(), '同意并继续')]"),
            (By.XPATH, "//button[text()='同意并继续']"),
            (By.CSS_SELECTOR, "button.primary"),
            (By.CSS_SELECTOR, ".primary"),
            (By.ID, "accept-button"),
            (By.XPATH, "//button[contains(@class, 'primary')]")
        ]
        
        consent_clicked = False
        
        # 首先尝试自动点击
        for locator in consent_button_locators:
            try:
                elements = driver.find_elements(*locator)
                for element in elements:
                    try:
                        text = element.text
                        if '同意' in text or '继续' in text:
                            print(f"找到按钮: '{text}'")
                            driver.execute_script("arguments[0].click();", element)
                            print("已点击'同意并继续'按钮")
                            consent_clicked = True
                            time.sleep(1)
                            break
                    except:
                        continue
                if consent_clicked:
                    break
            except:
                continue
        
        # 如果自动点击失败，提示用户手动点击
        if not consent_clicked:
            print("未能自动点击'同意并继续'按钮，请手动点击")
            input("点击后按Enter键继续...")
        
        # 等待页面过渡
        time.sleep(1)
        
        # ===== 步骤2: 输入邮箱 =====
        print("\n步骤2: 输入邮箱")
        
        # 方法1: 尝试定位邮箱输入框并输入（只输入用户名部分，不包含@和后缀）
        email_input_locators = [
            (By.ID, "MemberName"),
            (By.NAME, "MemberName"),
            (By.XPATH, "//input[@type='email']"),
            (By.CSS_SELECTOR, "input[type='email']")
        ]
        
        # 重要：只输入用户名部分，不含@和域名
        email_input_success = try_multiple_locators(
            driver, email_input_locators, "input", email_username, 5, 1, "邮箱输入框"
        )
        
        # 如果定位失败，尝试直接发送键盘输入
        if not email_input_success:
            print("未能定位邮箱输入框，尝试直接发送键盘输入...")
            actions.send_keys(email_username).perform()  # 只输入用户名部分
            print(f"已直接输入邮箱用户名: {email_username}")
            time.sleep(1)
        
        # 点击下一步按钮
        next_button_locators = [
            (By.ID, "iSignupAction"),
            (By.XPATH, "//button[@type='submit']"),
            (By.XPATH, "//input[@type='submit']"),
            (By.CSS_SELECTOR, "input[type='submit']"),
            (By.CSS_SELECTOR, "button[type='submit']")
        ]
        
        next_clicked = try_multiple_locators(
            driver, next_button_locators, "click", None, 5, 2, "下一步按钮"
        )
        
        # 如果点击失败，尝试使用Tab和Enter
        if not next_clicked:
            print("未能点击下一步按钮，尝试使用Tab+Enter键...")
            actions.send_keys(Keys.TAB).send_keys(Keys.ENTER).perform()
            print("已使用Tab+Enter进行下一步")
            time.sleep(2)
        
        # ===== 步骤3: 输入密码 =====
        print("\n步骤3: 设置密码")
        
        # 尝试定位密码输入框
        password_input_locators = [
            (By.ID, "PasswordInput"),
            (By.NAME, "Password"),
            (By.XPATH, "//input[@type='password']"),
            (By.CSS_SELECTOR, "input[type='password']")
        ]
        
        password_input_success = try_multiple_locators(
            driver, password_input_locators, "input", password, 5, 1, "密码输入框"
        )
        
        # 如果定位失败，尝试直接发送键盘输入
        if not password_input_success:
            print("未能定位密码输入框，尝试直接发送键盘输入...")
            actions.send_keys(password).perform()
            print("已直接输入密码")
            time.sleep(1)
        
        # 点击下一步按钮
        next_clicked = try_multiple_locators(
            driver, next_button_locators, "click", None, 5, 2, "下一步按钮"
        )
        
        # 如果点击失败，尝试使用Tab和Enter
        if not next_clicked:
            print("未能点击下一步按钮，尝试使用Tab+Enter键...")
            actions.send_keys(Keys.TAB).send_keys(Keys.ENTER).perform()
            print("已使用Tab+Enter进行下一步")
            time.sleep(2)
        
        # ===== 步骤4: 输入姓名 =====
        print("\n步骤4: 输入姓名")
        
        # 尝试定位名字输入框
        first_name_locators = [
            (By.ID, "FirstName"),
            (By.NAME, "FirstName"),
            (By.XPATH, "//input[contains(@placeholder, 'First')]")
        ]
        
        first_name_input_success = try_multiple_locators(
            driver, first_name_locators, "input", first_name, 5, 1, "名字输入框"
        )
        
        # 如果定位失败，尝试直接发送键盘输入
        if not first_name_input_success:
            print("未能定位名字输入框，尝试直接发送键盘输入...")
            actions.send_keys(first_name).perform()
            print(f"已直接输入名字: {first_name}")
            time.sleep(1)
        
        # 使用Tab键移动到姓氏输入框
        actions.send_keys(Keys.TAB).perform()
        time.sleep(1)
        
        # 输入姓氏
        last_name_locators = [
            (By.ID, "LastName"),
            (By.NAME, "LastName"),
            (By.XPATH, "//input[contains(@placeholder, 'Last')]")
        ]
        
        last_name_input_success = try_multiple_locators(
            driver, last_name_locators, "input", last_name, 5, 1, "姓氏输入框"
        )
        
        # 如果定位失败，尝试直接发送键盘输入
        if not last_name_input_success and not first_name_input_success:
            print("未能定位姓氏输入框，尝试直接发送键盘输入...")
            actions.send_keys(last_name).perform()
            print(f"已直接输入姓氏: {last_name}")
            time.sleep(1)
        
        # 点击下一步按钮
        next_clicked = try_multiple_locators(
            driver, next_button_locators, "click", None, 5, 2, "下一步按钮"
        )
        
        # 如果点击失败，尝试使用Tab和Enter
        if not next_clicked:
            print("未能点击下一步按钮，尝试使用Tab+Enter键...")
            actions.send_keys(Keys.TAB).send_keys(Keys.ENTER).perform()
            print("已使用Tab+Enter进行下一步")
            time.sleep(2)
        
        # ===== 步骤5: 设置生日 =====
        print("\n步骤5: 设置生日")
        
        # 1. 先输入年份
        year_input_success = False
        try:
            year_input = driver.find_element(By.ID, "BirthYear")
            year_input.clear()
            year_input.send_keys(str(birth_date['year']))
            year_input_success = True
            print(f"已输入年份: {birth_date['year']}")
            time.sleep(0.1)
        except:
            try:
                # 尝试通过Tab键定位年份输入框
                actions.send_keys(Keys.TAB).perform()
                time.sleep(0.1)
                actions.send_keys(str(birth_date['year'])).perform()
                year_input_success = True
                print(f"已使用键盘输入年份: {birth_date['year']}")
                time.sleep(0.1)
            except:
                print("未能输入年份")
        
        # 2. 移动到月份下拉框并选择
        # 尝试通过Tab键移动到月份
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.1)
        
        # 选择随机月份
        random_month = random.randint(1, 12)
        
        # 打开下拉列表
        actions.send_keys(Keys.SPACE).perform()
        time.sleep(0.1)
        
        # 先按HOME键回到第一个选项
        actions.send_keys(Keys.HOME).perform()
        time.sleep(0.1)
        
        # 按下键移动对应次数
        for _ in range(random_month):
            actions.send_keys(Keys.DOWN).perform()
            time.sleep(0.02)
        
        # 按Enter确认选择
        actions.send_keys(Keys.ENTER).perform()
        print(f"已选择月份: {random_month}")
        time.sleep(0.1)
        
        # 3. 移动到日期下拉框并选择
        # 通过Tab键移动到日期
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.1)
        
        # 选择随机日期
        random_day = random.randint(1, 28)
        
        # 打开下拉列表
        actions.send_keys(Keys.SPACE).perform()
        time.sleep(0.1)
        
        # 先按HOME键回到第一个选项
        actions.send_keys(Keys.HOME).perform()
        time.sleep(0.1)
        
        # 按下键移动对应次数
        for _ in range(random_day):
            actions.send_keys(Keys.DOWN).perform()
            time.sleep(0.02)
        
        # 按Enter确认选择
        actions.send_keys(Keys.ENTER).perform()
        print(f"已选择日期: {random_day}")
        time.sleep(0.1)
        
        # 直接按Enter提交表单
        actions.send_keys(Keys.ENTER).perform()
        time.sleep(0.3)
        
        # 如果直接提交失败，再尝试点击下一步按钮
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit'], button[type='submit']")
            driver.execute_script("arguments[0].click();", next_button)
            print("已点击下一步按钮")
            time.sleep(0.3)
        except:
            # 实在不行就用Tab+Enter
            actions.send_keys(Keys.TAB).send_keys(Keys.ENTER).perform()
            print("已使用Tab+Enter进行下一步")
            time.sleep(0.3)
        
        # ===== 步骤6: 选择国家 =====
        print("\n步骤6: 选择国家")
        
        # 我们不需要修改国家，可能已经默认选择了中国
        # 直接尝试点击下一步按钮
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit'], button[type='submit']")
            driver.execute_script("arguments[0].click();", next_button)
            print("已点击下一步按钮，跳过国家选择")
            time.sleep(0.3)
        except:
            try:
                # 如果找不到下一步按钮，可能需要先选择国家
                # 尝试通过Tab键移动到国家选择框
                actions.send_keys(Keys.TAB).perform()
                time.sleep(0.1)
                
                # 直接按Enter确认默认国家
                actions.send_keys(Keys.ENTER).perform()
                time.sleep(0.3)
                
                # 再次尝试点击下一步
                try:
                    next_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit'], button[type='submit']")
                    driver.execute_script("arguments[0].click();", next_button)
                    print("已点击下一步按钮")
                    time.sleep(0.3)
                except:
                    actions.send_keys(Keys.TAB).send_keys(Keys.ENTER).perform()
                    print("已使用Tab+Enter进行下一步")
                    time.sleep(0.3)
            except:
                print("无法进行国家选择，如需手动选择请操作")
                time.sleep(0.3)
        
        # ===== 步骤7: 人机验证或其他验证 =====
        print("\n步骤7: 可能需要人机验证")
        print("⚠️ 注意：如有验证码或人机验证，请手动完成...")
        
        # 提示用户进行后续操作
        print("\n注册过程基本完成，如需手动操作，请现在进行...")
        input("按Enter键继续，或完成其他必要的验证步骤...")
        
        # 最终确认
        print("\n注册流程已完成!")
        print("账号信息回顾:")
        print(f"邮箱: {email}")
        print(f"密码: {password}")
        
    except Exception as e:
        print(f"发生错误: {e}")
    finally:
        # 询问是否关闭浏览器
        close_browser = input("\n是否关闭浏览器? (y/n): ")
        if close_browser.lower() == 'y':
            driver.quit()
            print("浏览器已关闭")
        else:
            print("浏览器将保持打开状态，请手动关闭")

if __name__ == "__main__":
    print("======= Outlook邮箱自动注册工具 =======")
    print("此脚本将自动完成Outlook邮箱的注册流程")
    print("1. 自动模式 - 尝试自动完成所有步骤")
    print("2. 手动辅助模式 - 生成信息并指导您完成注册")
    
    mode = input("\n请选择模式 (默认1): ") or "1"
    
    try:
        if mode == "2":
            register_outlook_account(auto_mode=False)
        else:
            register_outlook_account(auto_mode=True)
        print("脚本执行完毕")
    except Exception as e:
        print(f"执行过程中发生错误: {e}")
        print("错误详情:")
        traceback.print_exc()
        input("按Enter键退出...") 