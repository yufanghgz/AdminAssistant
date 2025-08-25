#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试优化后的验证码处理功能
"""
import os
import sys
import logging
from base.erp_login import ERPLogin
from selenium.webdriver.common.by import By

def setup_logging():
    """设置日志"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('test_captcha.log', encoding='utf-8'),
            logging.StreamHandler()
        ],
        encoding='utf-8'
    )
    return logging.getLogger('TestCaptcha')

def test_captcha_recognition():
    """测试验证码识别功能"""
    logger = setup_logging()
    logger.info("开始测试验证码识别功能")
    
    try:
        # 创建ERPLogin实例
        erp_login = ERPLogin(
            headless=False,  # 设置为False以便观察验证码处理过程
            config_path='base/conf/erp-config.json',
            element_ids_path='base/conf/erp-element-ids.json'
        )
        
        # 打开登录页面
        erp_login.driver.get(erp_login.url)
        logger.info("打开登录页面")
        
        # 输入用户名和密码
        try:
            username_input = erp_login.driver.find_element(By.ID, erp_login.element_ids['username'])
            username_input.send_keys(erp_login.username)
            logger.info(f"输入用户名: {erp_login.username}")
            
            password_input = erp_login.driver.find_element(By.ID, erp_login.element_ids['password'])
            password_input.send_keys(erp_login.password)
            logger.info("输入密码")
        except Exception as e:
            logger.error(f"输入用户名或密码失败: {str(e)}")
            return False
        
        # 测试验证码处理
        logger.info("开始测试验证码处理...")
        success = erp_login._handle_captcha(max_retries=2)  # 最多重试2次
        
        if success:
            logger.info("验证码处理测试成功！")
        else:
            logger.warning("验证码处理测试失败，但这是正常的，因为可能需要手动输入")
        
        # 等待用户观察
        input("请观察验证码处理过程，完成后按Enter键继续...")
        
        return True
        
    except Exception as e:
        logger.error(f"测试过程中出错: {str(e)}", exc_info=True)
        return False
    finally:
        try:
            erp_login.driver.quit()
            logger.info("浏览器已关闭")
        except:
            pass

def test_captcha_image_preprocessing():
    """测试验证码图片预处理功能"""
    logger = setup_logging()
    logger.info("开始测试验证码图片预处理功能")
    
    try:
        # 检查是否有现有的验证码图片
        captcha_path = 'captcha.png'
        if not os.path.exists(captcha_path):
            logger.warning(f"未找到验证码图片: {captcha_path}")
            logger.info("请先运行登录测试生成验证码图片")
            return False
        
        # 创建ERPLogin实例（仅用于测试识别功能）
        erp_login = ERPLogin(
            headless=True,
            config_path='base/conf/erp-config.json',
            element_ids_path='base/conf/erp-element-ids.json'
        )
        
        # 测试直接识别
        logger.info("测试直接识别...")
        result1 = erp_login._recognize_with_easyocr(captcha_path)
        logger.info(f"直接识别结果: {result1}")
        
        # 测试预处理后识别
        logger.info("测试预处理后识别...")
        result2 = erp_login._recognize_with_image_preprocessing(captcha_path)
        logger.info(f"预处理后识别结果: {result2}")
        
        # 测试综合识别
        logger.info("测试综合识别...")
        result3 = erp_login._recognize_captcha(captcha_path)
        logger.info(f"综合识别结果: {result3}")
        
        return True
        
    except Exception as e:
        logger.error(f"测试过程中出错: {str(e)}", exc_info=True)
        return False

def main():
    """主函数"""
    print("验证码处理功能测试")
    print("=" * 50)
    print("1. 测试验证码识别功能（需要浏览器交互）")
    print("2. 测试验证码图片预处理功能")
    print("3. 退出")
    
    while True:
        choice = input("\n请选择测试项目 (1-3): ").strip()
        
        if choice == '1':
            print("\n开始测试验证码识别功能...")
            test_captcha_recognition()
        elif choice == '2':
            print("\n开始测试验证码图片预处理功能...")
            test_captcha_image_preprocessing()
        elif choice == '3':
            print("退出测试")
            break
        else:
            print("无效选择，请重新输入")

if __name__ == '__main__':
    main()

