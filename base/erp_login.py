#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ERP系统自动登录工具
用于登录 http://erp.ex.cosmosource.com:24080/csom/index.html
"""
import os
import json
import time
import logging
import base64
import io
from PIL import Image
import easyocr
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

class ERPLogin:
    def __init__(self, username=None, password=None, headless=False, config_path='base/conf/erp-config.json', element_ids_path='base/conf/erp-element-ids.json', element_ids=None):
        """
        初始化ERP登录器
        :param username: 用户名
        :param password: 密码
        :param headless: 是否以无头模式运行
        :param config_path: 配置文件路径
        :param element_ids_path: 元素ID配置文件路径
        :param element_ids: 页面元素ID字典，包含username, password, captcha, captchaImg, loginBtn, mainContent
        """
        # 配置日志
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('erp_login.log', encoding='utf-8'),
                logging.StreamHandler()
            ],
            encoding='utf-8'
        )
        self.logger = logging.getLogger('ERPLogin')

        # 初始化EasyOCR读取器（只初始化一次以提高性能）
        try:
            self.reader = easyocr.Reader(['en'], gpu=False)  # 使用英文识别，CPU模式
            self.logger.info("EasyOCR初始化成功")
        except Exception as e:
            self.logger.warning(f"EasyOCR初始化失败: {str(e)}")
            self.reader = None

        # 读取配置文件
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            self.logger.info(f"成功读取配置文件: {config_path}")
            self.username = username or config.get('username')
            self.password = password or config.get('password')
        except Exception as e:
            self.logger.error(f"读取配置文件失败: {str(e)}")
            self.username = username
            self.password = password

        self.url = 'http://erp.ex.cosmosource.com:24080/csom/index.html'
        self.headless = headless

        # 默认元素ID
        default_element_ids = {
            'username': 'username',
            'password': 'password',
            'captcha': 'captcha',
            'captchaImg': 'captchaImg',
            'loginBtn': 'loginBtn',
            'mainContent': 'mainContent'
        }

        # 读取元素ID配置文件
        try:
            if os.path.exists(element_ids_path):
                with open(element_ids_path, 'r', encoding='utf-8') as f:
                    file_element_ids = json.load(f)
                self.logger.info(f"成功读取元素ID配置文件: {element_ids_path}")
                # 合并默认值和配置文件中的值
                default_element_ids.update(file_element_ids)
            else:
                self.logger.warning(f"元素ID配置文件不存在: {element_ids_path}，使用默认值")
        except Exception as e:
            self.logger.error(f"读取元素ID配置文件失败: {str(e)}，使用默认值")

        # 设置元素ID，优先级: 参数 > 配置文件 > 默认值
        self.element_ids = {**default_element_ids, **(element_ids or {})}
        self.logger.info(f"最终使用的元素ID: {self.element_ids}")
        self.driver = self._init_driver(headless)
        self.logged_in = False

    def _init_driver(self, headless):
        """
        初始化浏览器驱动
        :param headless: 是否以无头模式运行
        :return: 浏览器驱动实例
        """
        chrome_options = Options()
        if headless:
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
        
        # 禁用自动化控制检测
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        try:
            # 使用webdriver_manager自动管理Chrome驱动
            driver = webdriver.Chrome(service=webdriver.ChromeService(ChromeDriverManager().install()), options=chrome_options)
            # 绕过webdriver检测
            driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
            })
            self.logger.info("浏览器驱动初始化成功")
            return driver
        except Exception as e:
            self.logger.error(f"浏览器驱动初始化失败: {str(e)}")
            raise

    def login(self, username=None, password=None):
        """
        登录ERP系统
        :param username: 用户名
        :param password: 密码
        :return: 是否登录成功
        """
        # 使用提供的参数或实例化时的参数
        self.username = username or self.username
        self.password = password or self.password

        if not all([self.username, self.password]):
            self.logger.error("用户名和密码不能为空")
            return False

        try:
            # 打开登录页面
            self.driver.get(self.url)
            self.logger.info(f"打开登录页面: {self.url}")
            time.sleep(3)  # 等待页面加载

            # 输入用户名
            try:
                username_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, self.element_ids['username']))
                )
                username_input.send_keys(self.username)
                self.logger.info(f"输入用户名: {self.username}")
            except Exception as e:
                self.logger.error(f"定位用户名输入框失败: {str(e)}")
                self.logger.error(f"使用的元素ID: {self.element_ids['username']}")
                return False

            # 输入密码
            try:
                password_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, self.element_ids['password']))
                )
                password_input.send_keys(self.password)
                self.logger.info("输入密码")
            except Exception as e:
                self.logger.error(f"定位密码输入框失败: {str(e)}")
                self.logger.error(f"使用的元素ID: {self.element_ids['password']}")
                return False

            # 处理验证码（包含登录按钮点击和验证）
            if not self._handle_captcha():
                return False

            # 验证是否登录成功
            if self._verify_login():
                self.logged_in = True
                self.logger.info("登录成功")
                return True
            else:
                self.logger.error("登录失败")
                return False
        except Exception as e:
            self.logger.error(f"登录过程中出错: {str(e)}")
            return False

    def _handle_captcha(self, max_retries=3):
        """
        处理图片验证码，支持重试机制
        :param max_retries: 最大重试次数
        :return: 是否成功处理
        """
        for attempt in range(max_retries):
            try:
                self.logger.info(f"验证码处理尝试 {attempt + 1}/{max_retries}")
                
                # 等待验证码图片加载
                try:
                    captcha_image = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, self.element_ids['captchaImg']))
                    )
                    self.logger.info("验证码图片加载完成")
                except Exception as e:
                    self.logger.error(f"定位验证码图片失败: {str(e)}")
                    self.logger.error(f"使用的元素ID: {self.element_ids['captchaImg']}")
                    return False

                # 尝试刷新验证码图片
                if attempt > 0:
                    try:
                        captcha_image.click()
                        self.logger.info("刷新验证码图片")
                        time.sleep(2)  # 等待新验证码加载
                    except Exception as e:
                        self.logger.warning(f"刷新验证码失败: {str(e)}")

                # 保存验证码图片到当前目录
                try:
                    captcha_path = os.path.join(os.getcwd(), f'captcha_attempt_{attempt + 1}.png')
                    captcha_image.screenshot(captcha_path)
                    self.logger.info(f"验证码图片已保存到: {captcha_path}")
                except Exception as e:
                    self.logger.warning(f"保存验证码图片失败: {str(e)}")
                    captcha_path = os.path.join(os.getcwd(), 'captcha.png')

                # 尝试自动识别验证码
                captcha_code = self._recognize_captcha(captcha_path)
                
                # 如果自动识别失败，提示用户手动输入
                if not captcha_code:
                    self.logger.warning(f"第{attempt + 1}次自动识别验证码失败，请手动输入")
                    captcha_code = input("请输入验证码: ")
                    if not captcha_code:
                        self.logger.error("验证码不能为空")
                        if attempt < max_retries - 1:
                            continue
                        return False

                # 输入验证码
                try:
                    captcha_input = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, self.element_ids['captcha']))
                    )
                    captcha_input.clear()  # 清除之前的内容
                    captcha_input.send_keys(captcha_code)
                    self.logger.info(f"输入验证码: {captcha_code}")
                    
                    # 点击登录按钮
                    try:
                        login_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[@class='btn' and contains(text(), '登')]"))
                        )
                        login_button.click()
                        self.logger.info("点击登录按钮")
                        time.sleep(3)  # 等待登录响应
                    except Exception as e:
                        self.logger.error(f"定位登录按钮失败: {str(e)}")
                        return False
                    
                    # 检查是否登录成功
                    if self._verify_login():
                        self.logger.info("验证码验证成功，登录成功")
                        return True
                    else:
                        # 检查是否有验证码错误提示
                        if self._check_captcha_error():
                            self.logger.warning(f"第{attempt + 1}次验证码错误")
                            if attempt < max_retries - 1:
                                # 清除验证码输入框，准备重试
                                captcha_input.clear()
                                continue
                            else:
                                self.logger.error("验证码重试次数已达上限")
                                return False
                        else:
                            self.logger.error("登录失败，但非验证码原因")
                            return False
                            
                except Exception as e:
                    self.logger.error(f"定位验证码输入框失败: {str(e)}")
                    self.logger.error(f"使用的元素ID: {self.element_ids['captcha']}")
                    return False
                    
            except Exception as e:
                self.logger.error(f"处理验证码时出错: {str(e)}")
                if attempt < max_retries - 1:
                    continue
                return False
        
        return False

    def _recognize_captcha(self, captcha_path):
        """
        使用EasyOCR识别验证码，支持多种识别策略
        :param captcha_path: 验证码图片路径
        :return: 识别出的验证码文本，失败返回None
        """
        try:
            # 检查EasyOCR是否可用
            if self.reader is None:
                self.logger.warning("EasyOCR未初始化，跳过自动识别")
                return None
            
            # 尝试多种识别策略
            strategies = [
                self._recognize_with_easyocr,
                self._recognize_with_image_preprocessing
            ]
            
            for strategy in strategies:
                try:
                    result = strategy(captcha_path)
                    if result:
                        return result
                except Exception as e:
                    self.logger.debug(f"识别策略失败: {str(e)}")
                    continue
            
            self.logger.warning("所有识别策略都失败了")
            return None
                
        except Exception as e:
            self.logger.error(f"验证码识别失败: {str(e)}")
            return None

    def _recognize_with_easyocr(self, captcha_path):
        """
        使用EasyOCR直接识别验证码
        """
        # 使用EasyOCR识别验证码
        results = self.reader.readtext(captcha_path)
        
        if not results:
            self.logger.warning("EasyOCR未识别到任何文本")
            return None
        
        # 提取所有识别到的文本
        detected_texts = []
        for (bbox, text, confidence) in results:
            # 清理文本，只保留字母和数字
            cleaned_text = ''.join(c for c in text if c.isalnum())
            if cleaned_text:
                detected_texts.append(cleaned_text)
                self.logger.debug(f"识别到文本: '{text}' -> '{cleaned_text}' (置信度: {confidence:.2f})")
        
        # 合并所有识别到的文本
        if detected_texts:
            captcha_text = ''.join(detected_texts)
            captcha_text = captcha_text.upper()  # 转换为大写
            
            # 验证识别结果是否合理（通常是4-6位字符）
            if len(captcha_text) >= 3 and len(captcha_text) <= 8:
                self.logger.info(f"EasyOCR识别验证码成功: {captcha_text}")
                return captcha_text
            else:
                self.logger.warning(f"验证码识别结果不合理: {captcha_text}")
                return None
        else:
            self.logger.warning("未识别到有效的验证码文本")
            return None

    def _recognize_with_image_preprocessing(self, captcha_path):
        """
        使用图像预处理后识别验证码
        """
        try:
            from PIL import Image, ImageEnhance, ImageFilter
            import numpy as np
            
            # 打开图片
            img = Image.open(captcha_path)
            
            # 转换为灰度图
            img_gray = img.convert('L')
            
            # 增强对比度
            enhancer = ImageEnhance.Contrast(img_gray)
            img_enhanced = enhancer.enhance(2.0)
            
            # 应用锐化滤镜
            img_sharp = img_enhanced.filter(ImageFilter.SHARPEN)
            
            # 保存预处理后的图片
            processed_path = captcha_path.replace('.png', '_processed.png')
            img_sharp.save(processed_path)
            
            # 使用预处理后的图片进行识别
            results = self.reader.readtext(processed_path)
            
            if not results:
                return None
            
            # 提取文本
            detected_texts = []
            for (bbox, text, confidence) in results:
                cleaned_text = ''.join(c for c in text if c.isalnum())
                if cleaned_text:
                    detected_texts.append(cleaned_text)
            
            if detected_texts:
                captcha_text = ''.join(detected_texts).upper()
                if len(captcha_text) >= 3 and len(captcha_text) <= 8:
                    self.logger.info(f"预处理后识别验证码成功: {captcha_text}")
                    return captcha_text
            
            return None
            
        except Exception as e:
            self.logger.debug(f"图像预处理识别失败: {str(e)}")
            return None

    def _verify_login(self):
        """
        验证是否登录成功
        :return: 是否登录成功
        """
        # 检查是否跳转到了首页，或者是否存在某个登录后的元素
        try:
            # 等待页面跳转或出现登录成功标志
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "content"))
            )
            self.logger.info("找到登录成功标志元素: content")
            return True
        except Exception as e:
            self.logger.error(f"验证登录成功失败: {str(e)}")
            return False

    def _check_captcha_error(self):
        """
        检查是否有验证码错误提示
        :return: 是否有验证码错误
        """
        try:
            # 等待一段时间，让错误信息显示
            time.sleep(2)
            
            # 检查常见的验证码错误提示
            error_selectors = [
                "//div[contains(text(), '验证码') and contains(text(), '错误')]",
                "//div[contains(text(), '验证码') and contains(text(), '不正确')]",
                "//span[contains(text(), '验证码') and contains(text(), '错误')]",
                "//span[contains(text(), '验证码') and contains(text(), '不正确')]",
                "//div[contains(@class, 'error') and contains(text(), '验证码')]",
                "//span[contains(@class, 'error') and contains(text(), '验证码')]"
            ]
            
            for selector in error_selectors:
                try:
                    error_element = self.driver.find_element(By.XPATH, selector)
                    if error_element.is_displayed():
                        self.logger.info(f"检测到验证码错误提示: {error_element.text}")
                        return True
                except:
                    continue
            
            return False
        except Exception as e:
            self.logger.debug(f"检查验证码错误时出错: {str(e)}")
            return False

    def logout(self):
        """
        退出登录
        """
        if self.logged_in:
            try:
                # 这里需要根据实际的登出按钮位置进行修改
                # logout_button = self.driver.find_element(By.ID, 'logoutBtn')
                # logout_button.click()
                self.logger.info("已退出登录")
                self.logged_in = False
            except Exception as e:
                self.logger.error(f"退出登录时出错: {str(e)}")
        self.driver.quit()
        self.logger.info("浏览器已关闭")

    def perform_operation(self, operation_name, **kwargs):
        """
        执行登录后的操作
        :param operation_name: 操作名称
        :param kwargs: 操作参数
        :return: 操作结果
        """
        if not self.logged_in:
            self.logger.error("未登录，无法执行操作")
            return False

        try:
            self.logger.info(f"执行操作: {operation_name}")
            # 这里根据不同的操作名称实现相应的功能
            # 例如，导航到某个页面，填写表单等
            # 示例：
            if operation_name == 'navigate_to_page':
                page_url = kwargs.get('page_url')
                if page_url:
                    self.driver.get(page_url)
                    time.sleep(3)
                    self.logger.info(f"导航到页面: {page_url}")
                    return True
                else:
                    self.logger.error("未提供页面URL")
                    return False
            # 可以添加更多操作...
            else:
                self.logger.error(f"未知操作: {operation_name}")
                return False
        except Exception as e:
            self.logger.error(f"执行操作时出错: {str(e)}")
            return False


def main():
    """
    主函数，用于演示如何使用ERPLogin类
    """
    # 创建日志记录器
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('erp_login.log', encoding='utf-8'),
            logging.StreamHandler()
        ],
        encoding='utf-8'
    )
    logger = logging.getLogger('ERPLoginMain')

    try:
        logger.info("开始ERP登录演示")
        
        # 创建ERPLogin实例
        # 不需要手动读取配置文件，ERPLogin类会自动处理
        erp_login = ERPLogin(
            headless=False,
            config_path='base/conf/erp-config.json',
            element_ids_path='base/conf/erp-element-ids.json'
        )
        
        # 登录ERP系统
        if erp_login.login():
            logger.info("ERP登录成功")
            # 登录成功后执行操作
            erp_login.perform_operation('navigate_to_page', page_url='http://erp.ex.cosmosource.com:24080/csom/index.html')
            
            # 保持浏览器打开一段时间
            input("操作完成后按Enter键退出...")
        else:
            logger.error("ERP登录失败")
    except Exception as e:
        logger.error(f"程序执行过程中出错: {str(e)}", exc_info=True)
    finally:
        # 退出登录
        try:
            erp_login.logout()
        except:
            pass


if __name__ == '__main__':
    main()