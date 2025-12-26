# 贡献指南

感谢您对Microsoft Graph MCP Server项目的关注！我们欢迎各种形式的贡献。

## 开发环境设置

1. 克隆项目仓库
2. 创建虚拟环境：
   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/Mac
   # 或
   venv\Scripts\activate  # Windows
   ```
3. 安装依赖：
   ```bash
   pip install -r requirements.txt
   pip install -e .
   ```

## 代码规范

- 使用Black进行代码格式化
- 使用isort进行导入排序
- 遵循PEP 8编码规范
- 为所有公共API编写文档字符串

## 提交代码

1. 创建功能分支
2. 实现功能并添加测试
3. 运行测试确保通过
4. 提交代码并创建Pull Request

## 测试

运行测试：
```bash
pytest
```

## 问题报告

请在GitHub Issues中报告问题，包括：
- 问题描述
- 复现步骤
- 期望行为
- 实际行为
- 环境信息