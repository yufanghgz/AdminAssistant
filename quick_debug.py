#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速调试ERP登录页面元素
"""
import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def quick_debug():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger('QuickDebug')
    
    # 初始化浏览器
    chrome_options = Options()
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    
    try:
        driver = webdriver.Chrome(service=webdriver.ChromeService(ChromeDriverManager().install()), options=chrome_options)
        logger.info("浏览器启动成功")
        
        # 打开登录页面
        driver.get('http://erp.ex.cosmosource.com:24080/csom/index.html')
        logger.info("页面加载中...")
        time.sleep(3)
        
        # 查找所有输入框
        inputs = driver.find_elements(By.TAG_NAME, "input")
        logger.info(f"找到 {len(inputs)} 个输入框:")
        
        for i, inp in enumerate(inputs):
            elem_id = inp.get_attribute('id') or '无ID'
            elem_name = inp.get_attribute('name') or '无name'
            elem_type = inp.get_attribute('type') or '无type'
            elem_placeholder = inp.get_attribute('placeholder') or '无placeholder'
            logger.info(f"  输入框{i+1}: ID='{elem_id}', name='{elem_name}', type='{elem_type}', placeholder='{elem_placeholder}'")
        
        # 查找所有按钮
        buttons = driver.find_elements(By.TAG_NAME, "button")
        logger.info(f"\n找到 {len(buttons)} 个按钮:")
        
        for i, btn in enumerate(buttons):
            elem_id = btn.get_attribute('id') or '无ID'
            elem_text = btn.text or '无文本'
            elem_class = btn.get_attribute('class') or '无class'
            logger.info(f"  按钮{i+1}: ID='{elem_id}', text='{elem_text}', class='{elem_class}'")
        
        # 查找所有图片
        images = driver.find_elements(By.TAG_NAME, "img")
        logger.info(f"\n找到 {len(images)} 个图片:")
        
        for i, img in enumerate(images):
            elem_id = img.get_attribute('id') or '无ID'
            elem_src = img.get_attribute('src') or '无src'
            elem_alt = img.get_attribute('alt') or '无alt'
            logger.info(f"  图片{i+1}: ID='{elem_id}', src='{elem_src}', alt='{elem_alt}'")
        
        # 查找可能的登录成功标志元素
        logger.info(f"\n查找可能的登录成功标志元素:")
        possible_main_elements = driver.find_elements(By.CSS_SELECTOR, "[id*='main'], [class*='main'], [id*='content'], [class*='content']")
        for i, elem in enumerate(possible_main_elements):
            elem_id = elem.get_attribute('id') or '无ID'
            elem_class = elem.get_attribute('class') or '无class'
            elem_tag = elem.tag_name
            logger.info(f"  元素{i+1}: tag='{elem_tag}', ID='{elem_id}', class='{elem_class}'")
        
        logger.info("\n=== 推荐的元素ID配置 ===")
        logger.info("请根据上面的输出更新 base/conf/erp-element-ids.json 文件")
        
    except Exception as e:
        logger.error(f"调试出错: {str(e)}")
    finally:
        try:
            driver.quit()
            logger.info("浏览器已关闭")
        except:
            pass

if __name__ == '__main__':
    quick_debug()

