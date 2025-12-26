# Microsoft Graph MCP Server

一个基于Microsoft Graph API的MCP (Model Context Protocol) Server，提供对Microsoft 365生态系统的全面访问。

## 功能特性

### 用户和组织管理
- 获取用户个人资料、照片、邮箱地址等基本信息
- 查询组织中的用户列表和组信息
- 管理用户权限和组成员身份

### 邮件和日历操作
- 读取和发送Outlook邮件
- 管理日历事件，创建、更新和删除会议
- 查询联系人和个人联系人信息

### 文件和文档管理
- 访问OneDrive和SharePoint中的文件和文件夹
- 上传、下载、移动和删除文件
- 管理文件权限和共享设置

### 团队协作功能
- 访问Teams团队和频道
- 管理团队成员和频道消息
- 创建和管理Planner任务和待办事项

### 数据分析和智能洞察
- 获取用户活动数据和常用文件趋势
- 分析会议请求和协作模式
- 生成个性化的数据洞察和决策支持

## 安装

```bash
pip install -r requirements.txt
```

## 配置

### 交互式认证（推荐）

本服务器使用设备代码流进行交互式认证，无需Azure应用注册或客户端密钥。

首次运行时，您将看到一个设备代码，需要：
1. 打开浏览器访问 `https://microsoft.com/devicelogin`
2. 输入显示的设备代码
3. 使用您的Microsoft账户登录

### 可选配置

如需使用自定义Azure应用注册，可创建 `.env` 文件：

```env
CLIENT_ID=your_custom_client_id
TENANT_ID=organizations
```

**注意**: 默认使用Microsoft的公共客户端ID，无需配置即可使用。

## 使用

### 在Claude Desktop中配置

编辑Claude Desktop配置文件：
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

添加以下配置：

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "python",
      "args": ["-m", "microsoft_graph_mcp_server.main"]
    }
  }
}
```

**注意**: 默认使用Microsoft的公共客户端ID和"organizations"租户，无需额外配置。

### 使用步骤

1. **首先运行 `login` 工具** - 这将触发设备代码流认证
2. **按照提示完成认证** - 打开浏览器并输入设备代码
3. **使用其他工具** - 认证成功后，所有工具都可以正常使用

### 可用工具

- **login** - 使用设备代码流进行Microsoft Graph认证（首先运行此工具）
- **get_user_info** - 获取当前用户信息
- **list_users** - 列出组织中的用户
- **get_messages** - 获取邮件
- **send_message** - 发送邮件
- **get_events** - 获取日历事件
- **create_event** - 创建日历事件
- **list_files** - 列出OneDrive文件和文件夹
- **get_teams** - 获取Microsoft Teams列表
- **get_team_channels** - 获取特定团队的频道

### 直接运行

```bash
# 启动MCP Server
python -m microsoft_graph_mcp_server.main

# 或使用安装后的命令
microsoft-graph-mcp-server
```

## 开发

```bash
# 安装开发依赖
pip install -e .

# 运行测试
pytest

# 代码格式化
black .
isort .
```

## 许可证

MIT License