# ERP系统自动登录工具

这是一个基于Selenium的ERP系统自动登录工具，旨在简化ERP系统的登录过程，并提供基础框架以便扩展其他操作。

## 功能特点
- 自动读取配置文件中的用户名和密码
- 支持自定义页面元素ID配置
- 自动处理验证码（当前为手动输入）
- 提供登录状态验证
- 支持无头模式运行
- 详细的日志记录

## 目录结构
```
AdminAssistent/
├── base/
│   ├── erp_login.py        # ERP登录核心类
│   └── conf/
│       ├── erp-config.json       # 登录凭据配置
│       └── erp-element-ids.json  # 页面元素ID配置
├── test_erp_login.py       # 测试脚本
└── README.md               # 使用说明
```

## 安装依赖
```bash
pip install selenium webdriver-manager
```

## 配置说明
1. **登录凭据配置**：编辑 `base/conf/erp-config.json` 文件，填入ERP系统的用户名和密码：
   ```json
   {
       "username": "your_username",
       "password": "your_password"
   }
   ```

2. **页面元素ID配置**：编辑 `base/conf/erp-element-ids.json` 文件，根据实际ERP系统页面元素ID进行调整：
   ```json
   {
       "username": "username_input_id",
       "password": "password_input_id",
       "captcha": "captcha_input_id",
       "captchaImg": "captcha_image_id",
       "loginBtn": "login_button_id",
       "mainContent": "main_content_id"
   }
   ```

## 使用方法
1. 确保已安装所有依赖
2. 配置好登录凭据和页面元素ID
3. 运行测试脚本：
   ```bash
   python test_erp_login.py
   ```

## 扩展功能
要添加登录后的自定义操作，可以继承 `ERPLogin` 类并重写 `perform_operation` 方法，或者直接在测试脚本中调用该方法后添加自定义代码。

## 注意事项
1. 本工具需要Chrome浏览器，请确保已安装最新版本的Chrome
2. 首次运行时，webdriver-manager会自动下载并安装匹配的ChromeDriver
3. 验证码当前需要手动输入，未来可考虑集成OCR服务实现自动化
4. 如果ERP系统页面结构发生变化，可能需要更新页面元素ID配置
5. 无头模式可能会被某些网站的反爬机制检测到，建议在测试环境下使用

## 示例代码
```python
from base.erp_login import ERPLogin

# 实例化ERPLogin对象
erp_login = ERPLogin(
    headless=False,
    config_path='base/conf/erp-config.json',
    element_ids_path='base/conf/erp-element-ids.json'
)

# 执行登录
success = erp_login.login()

if success:
    print("登录成功")
    # 执行自定义操作
    erp_login.perform_operation()
else:
    print("登录失败")

# 关闭浏览器
erp_login.close()
```