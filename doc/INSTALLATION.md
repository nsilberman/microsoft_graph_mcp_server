# 安装和配置指南

## 前置要求

- Python 3.8 或更高版本
- Microsoft账户（个人账户或工作/学校账户）

## 步骤1：安装项目

```bash
# 克隆项目
git clone <repository-url>
cd microsoft_graph_mcp_server

# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
pip install -e .
```

## 步骤2：配置Claude Desktop

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

## 步骤3：首次使用 - 交互式认证

### 使用login工具进行认证

1. 在Claude Desktop中，首先调用 **`login`** 工具
2. 服务器将显示设备代码认证提示：

```
======================================================================
MICROSOFT GRAPH AUTHENTICATION
======================================================================

To sign in, use a web browser to open the page:

https://microsoft.com/devicelogin

And enter the code:

ABCD-EFGH-IJKL

======================================================================
Waiting for authentication...
======================================================================
```

3. 打开浏览器访问 `https://microsoft.com/devicelogin`
4. 输入显示的设备代码
5. 使用您的Microsoft账户登录
6. 授予必要的权限
7. 认证成功后，login工具将返回成功消息

### 使用其他工具

认证成功后，您可以使用以下工具：
- **get_user_info** - 获取当前用户信息
- **list_users** - 列出组织中的用户
- **get_messages** - 获取邮件
- **get_events** - 获取日历事件
- **create_event** - 创建日历事件
- **list_files** - 列出OneDrive文件和文件夹
- **get_teams** - 获取Microsoft Teams列表
- **get_team_channels** - 获取特定团队的频道

**重要**: 如果您在未认证的情况下调用其他工具，它们将自动触发认证流程，但建议先使用login工具以确保认证成功。

## 可选配置

### 使用自定义Azure应用注册

如果您有自己的Azure应用注册，可以创建 `.env` 文件：

```bash
cp .env.example .env
```

编辑 `.env` 文件：

```env
CLIENT_ID=your_custom_client_id
TENANT_ID=organizations
```

### 环境变量配置

您也可以通过环境变量配置：

```bash
# Windows PowerShell
$env:CLIENT_ID="your_client_id"
$env:TENANT_ID="organizations"

# Linux/Mac
export CLIENT_ID="your_client_id"
export TENANT_ID="organizations"
```

## 运行服务器

```bash
# 使用Python模块运行
python -m microsoft_graph_mcp_server.main

# 或使用安装后的命令
microsoft-graph-mcp-server
```

## 故障排除

### 常见问题

1. **认证失败**：
   - 确保输入的设备代码正确
   - 检查网络连接
   - 确认您的Microsoft账户有权限访问所需资源

2. **权限不足**：
   - 确保在认证时授予了所有必要的权限
   - 某些功能可能需要管理员批准

3. **连接超时**：
   - 检查网络连接和防火墙设置
   - 设备代码有15分钟有效期，超时需要重新认证

### 调试模式

启用详细日志：

```bash
python -m microsoft_graph_mcp_server.main --verbose
```

### 重新认证

如需更换账户或重新认证，删除缓存令牌并重新启动服务器即可。

## 安全说明

- 使用设备代码流时，无需存储客户端密钥
- 访问令牌会自动刷新，无需重新认证
- 令牌缓存在本地，确保您的计算机安全
- 如需清除认证信息，删除本地缓存文件