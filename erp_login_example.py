#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ERP系统登录示例
"""
import logging
from base.erp_login import ERPLogin

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('erp_login_example.log', encoding='utf-8'),
        logging.StreamHandler()
    ],
    encoding='utf-8'
)
logger = logging.getLogger('ERPLoginExample')


def main():
    """
    主函数，演示ERP系统登录和基本操作
    """
    logger.info("开始ERP系统登录示例")

    try:
        # 创建ERPLogin实例
        # 可以直接传入用户名和密码，或者留空从配置文件读取
        # erp = ERPLogin(username='your_username', password='your_password')
        erp = ERPLogin()

        # 登录ERP系统
        if erp.login():
            logger.info("登录成功，可以开始执行操作")

            # 示例操作1: 导航到某个页面
            # 请替换为实际的ERP系统页面URL
            # erp.perform_operation('navigate_to_page', 
            #                      page_url='http://erp.ex.cosmosource.com:24080/csom/somepage')

            # 示例操作2: 这里可以添加更多自定义操作
            # 例如，查询数据、提交表单等

            # 保持程序运行，以便查看结果
            input("操作完成后按Enter键退出...")
        else:
            logger.error("登录失败，无法继续")
    except Exception as e:
        logger.error(f"程序执行过程中出错: {str(e)}")
    finally:
        # 确保退出登录
        try:
            erp.logout()
        except:
            pass

    logger.info("ERP系统登录示例结束")


if __name__ == '__main__':
    main()