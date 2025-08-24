#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ERP登录页面调试工具
用于检查页面元素结构
"""
import os
import json
import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def debug_erp_page():
    """
    调试ERP登录页面，检查元素结构
    """
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger('ERPDebug')

    # 初始化浏览器
    chrome_options = Options()
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    try:
        driver = webdriver.Chrome(service=webdriver.ChromeService(ChromeDriverManager().install()), options=chrome_options)
        driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
        })
        logger.info("浏览器驱动初始化成功")
    except Exception as e:
        logger.error(f"浏览器驱动初始化失败: {str(e)}")
        return

    try:
        # 打开登录页面
        url = 'http://erp.ex.cosmosource.com:24080/csom/index.html'
        driver.get(url)
        logger.info(f"打开登录页面: {url}")
        
        # 等待页面加载
        time.sleep(5)
        
        # 检查页面标题
        logger.info(f"页面标题: {driver.title}")
        
        # 查找所有输入框
        logger.info("=== 查找所有输入框 ===")
        input_elements = driver.find_elements(By.TAG_NAME, "input")
        for i, elem in enumerate(input_elements):
            try:
                elem_id = elem.get_attribute('id') or '无ID'
                elem_name = elem.get_attribute('name') or '无name'
                elem_type = elem.get_attribute('type') or '无type'
                elem_placeholder = elem.get_attribute('placeholder') or '无placeholder'
                elem_class = elem.get_attribute('class') or '无class'
                
                logger.info(f"输入框 {i+1}: ID='{elem_id}', name='{elem_name}', type='{elem_type}', placeholder='{elem_placeholder}', class='{elem_class}'")
            except Exception as e:
                logger.error(f"获取输入框 {i+1} 属性失败: {str(e)}")
        
        # 查找所有按钮
        logger.info("=== 查找所有按钮 ===")
        button_elements = driver.find_elements(By.TAG_NAME, "button")
        for i, elem in enumerate(button_elements):
            try:
                elem_id = elem.get_attribute('id') or '无ID'
                elem_text = elem.text or '无文本'
                elem_class = elem.get_attribute('class') or '无class'
                
                logger.info(f"按钮 {i+1}: ID='{elem_id}', text='{elem_text}', class='{elem_class}'")
            except Exception as e:
                logger.error(f"获取按钮 {i+1} 属性失败: {str(e)}")
        
        # 查找所有图片
        logger.info("=== 查找所有图片 ===")
        img_elements = driver.find_elements(By.TAG_NAME, "img")
        for i, elem in enumerate(img_elements):
            try:
                elem_id = elem.get_attribute('id') or '无ID'
                elem_src = elem.get_attribute('src') or '无src'
                elem_alt = elem.get_attribute('alt') or '无alt'
                elem_class = elem.get_attribute('class') or '无class'
                
                logger.info(f"图片 {i+1}: ID='{elem_id}', src='{elem_src}', alt='{elem_alt}', class='{elem_class}'")
            except Exception as e:
                logger.error(f"获取图片 {i+1} 属性失败: {str(e)}")
        
        # 尝试查找可能的用户名输入框
        logger.info("=== 尝试查找用户名输入框 ===")
        possible_username_selectors = [
            "input[type='text']",
            "input[placeholder*='用户名']",
            "input[placeholder*='user']",
            "input[placeholder*='账号']",
            "input[name*='username']",
            "input[name*='user']",
            "input[id*='username']",
            "input[id*='user']"
        ]
        
        for selector in possible_username_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                if elements:
                    logger.info(f"找到匹配选择器 '{selector}' 的元素: {len(elements)} 个")
                    for i, elem in enumerate(elements):
                        elem_id = elem.get_attribute('id') or '无ID'
                        elem_name = elem.get_attribute('name') or '无name'
                        elem_placeholder = elem.get_attribute('placeholder') or '无placeholder'
                        logger.info(f"  - 元素 {i+1}: ID='{elem_id}', name='{elem_name}', placeholder='{elem_placeholder}'")
            except Exception as e:
                logger.error(f"使用选择器 '{selector}' 查找失败: {str(e)}")
        
        # 尝试查找可能的密码输入框
        logger.info("=== 尝试查找密码输入框 ===")
        possible_password_selectors = [
            "input[type='password']",
            "input[placeholder*='密码']",
            "input[placeholder*='password']",
            "input[name*='password']",
            "input[name*='pwd']",
            "input[id*='password']",
            "input[id*='pwd']"
        ]
        
        for selector in possible_password_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                if elements:
                    logger.info(f"找到匹配选择器 '{selector}' 的元素: {len(elements)} 个")
                    for i, elem in enumerate(elements):
                        elem_id = elem.get_attribute('id') or '无ID'
                        elem_name = elem.get_attribute('name') or '无name'
                        elem_placeholder = elem.get_attribute('placeholder') or '无placeholder'
                        logger.info(f"  - 元素 {i+1}: ID='{elem_id}', name='{elem_name}', placeholder='{elem_placeholder}'")
            except Exception as e:
                logger.error(f"使用选择器 '{selector}' 查找失败: {str(e)}")
        
        # 保存页面源码
        page_source = driver.page_source
        with open('erp_page_source.html', 'w', encoding='utf-8') as f:
            f.write(page_source)
        logger.info("页面源码已保存到 erp_page_source.html")
        
        # 截图
        driver.save_screenshot('erp_page_screenshot.png')
        logger.info("页面截图已保存到 erp_page_screenshot.png")
        
        input("按Enter键关闭浏览器...")
        
    except Exception as e:
        logger.error(f"调试过程中出错: {str(e)}")
    finally:
        driver.quit()
        logger.info("浏览器已关闭")

if __name__ == '__main__':
    debug_erp_page()

