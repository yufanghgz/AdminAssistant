#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ERP登录测试脚本
用于测试ERPLogin类的功能
"""
from base.erp_login import ERPLogin
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('ERPLoginTest')

def main():
    try:
        logger.info("开始ERP登录测试")
        
        # 实例化ERPLogin对象
        # 这里可以不传入用户名和密码，它们会从配置文件中读取
        # 也可以不传入element_ids，它们会从erp-element-ids.json中读取
        erp_login = ERPLogin(
            headless=False,
            config_path='base/conf/erp-config.json',
            element_ids_path='base/conf/erp-element-ids.json'
        )
        
        # 执行登录
        logger.info("开始登录ERP系统")
        success = erp_login.login()
        
        if success:
            logger.info("ERP系统登录成功")
            # 这里可以添加后续操作
            erp_login.perform_operation()
        else:
            logger.error("ERP系统登录失败")
        
        # 关闭浏览器
        erp_login.close()
        logger.info("测试完成")
    except Exception as e:
        logger.error(f"测试过程中发生错误: {str(e)}", exc_info=True)

if __name__ == '__main__':
    main()