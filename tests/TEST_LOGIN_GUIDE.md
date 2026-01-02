# Login 和 Status Check 测试说明

## 测试脚本

运行以下命令来测试登录和状态检查功能：

```bash
cd C:/Project/microsoft_graph_mcp_server
python test_login_complete.py
```

## 测试脚本功能

`test_login_complete.py` 脚本会：

1. **检查初始登录状态** - 确认当前未认证
2. **发起登录** - 获取设备代码和验证 URL
3. **使用 device_code 检查状态** - 验证设备流程是否正确保存
4. **不使用 device_code 检查状态** - 测试常规状态检查
5. **显示日志文件内容** - 查看所有生成的日志
6. **显示令牌文件内容** - 查看保存的认证数据

## 日志文件位置

- **主日志**: `microsoft_graph_mcp_server/mcp_server_auth.log`
- **令牌文件**: `~/.microsoft_graph_mcp_tokens.json`
- **设备流程文件**: `~/.microsoft_graph_mcp_device_flows.json`

## 日志格式

### 启动日志

每次服务器启动时，您会看到：

```
======================================================================
MCP SERVER STARTED
Timestamp: 1767316842.1134696
Log file: C:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\cp_server_auth.log
======================================================================

2026-01-02 09:28:26,873 - root - INFO - ======================================================================
2026-01-02 09:28:26,873 - root - INFO - Logging initialized successfully
2026-01-02 09:28:26,873 - root - INFO - Log file: C:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\cp_server_auth.log
2026-01-02 09:28:26,873 - root - INFO - Server: microsoft-graph-mcp-server v0.1.0
2026-01-02 09:28:26,873 - root - INFO - ============================================================
```

### 运行时日志

当您调用 auth 工具时，您会看到类似以下的日志：

```
2026-01-02 09:30:00,123 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Initiating device code flow (attempt 1/3)
2026-01-02 09:30:00,456 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Device code flow initiated successfully. user_code: ABC123, device_code: 0.ARoA6Wg...
2026-01-02 09:30:05,789 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - check_login_status called with device_code: 0.AroA6Wg...
2026-01-02 09:30:06,123 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Token loaded from disk - authenticated: True, has_token: True
```

## 完整的认证流程

### 1. 调用 login

在 MCP 客户端中调用：
```json
{
  "action": "login"
}
```

**响应**：
```json
{
  "status": "pending",
  "message": "Please complete authentication using the link and code below...",
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345",
  "device_code": "0.ARoA6Wg...",
  "expires_in": 900,
  "interval": 5
}
```

**控制台输出**：
```
======================================================================
MICROSOFT GRAPH AUTHENTICATION
======================================================================

To sign in, use a web browser to open page:

https://microsoft.com/devicelogin

And enter code:

ABC12345

======================================================================
Please complete authentication in your browser, then call check_status.
======================================================================
```

**日志**：
```
2026-01-02 09:30:00,123 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Initiating device code flow (attempt 1/3)
2026-01-02 09:30:00,456 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Device code flow initiated successfully. user_code: ABC12345, device_code: 0.AroA6Wg...
```

### 2. 在浏览器中完成认证

1. 打开 `verification_uri`
2. 输入 `user_code`
3. 登录您的 Microsoft 账户
4. 授予权限

### 3. 调用 check_status（带 device_code）

```json
{
  "action": "check_status",
  "device_code": "0.ARoA6Wg..."
}
```

**响应（等待前）**：
```json
{
  "status": "pending",
  "message": "Authentication still pending. Please complete authentication in the browser and call check_status again.",
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345"
}
```

**响应（认证成功后）**：
```json
{
  "status": "authenticated",
  "message": "Authenticated with Microsoft Graph. Token expires in 59 minutes and 59 seconds at 2026-01-02 10:30:00",
  "token_expiry": 1735216200,
  "expiry_datetime": "2026-01-02 10:30:00",
  "remaining_seconds": 3599,
  "remaining_minutes": 59,
  "remaining_hours": 0
}
```

**日志**：
```
2026-01-02 09:30:05,789 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - check_login_status called with device_code: 0.AroA6Wg...
2026-01-02 09:30:05,790 - microsoft_graph_mcp_server.auth_modules.token_manager - INFO - Loading tokens from disk
2026-01-02 09:30:06,123 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Processing device flow with device_code: 0.AroA6Wg...
2026-01-02 09:30:06,456 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - acquire_token_by_device_flow result error: none
2026-01-02 09:30:06,789 - microsoft_graph_mcp_server.auth_modules.token_manager - INFO - Tokens saved to disk successfully
```

### 4. 调用 check_status（不带 device_code）

```json
{
  "action": "check_status"
}
```

**响应**：
```json
{
  "status": "authenticated",
  "message": "Authenticated with Microsoft Graph. Token expires in 45 minutes at 2026-01-02 10:15:00",
  "token_expiry": 1735214900,
  "expiry_datetime": "2026-01-02 10:15:00",
  "remaining_seconds": 2700,
  "remaining_minutes": 45,
  "remaining_hours": 0
}
```

## 令牌文件格式

### token_file (~/.microsoft_graph_mcp_tokens.json)

```json
{
  "access_token": "eyJ0eXAiOiJKV1QiLCJub25jZSI6...",
  "refresh_token": "0.ARoA6Wg...",
  "token_expiry": 1735216200.123,
  "authenticated": true
}
```

### device_flow_file (~/.microsoft_graph_mcp_device_flows.json)

```json
{
  "0.ARoA6Wg...": {
    "user_code": "ABC12345",
    "device_code": "0.ARoA6Wg...",
    "verification_uri": "https://microsoft.com/devicelogin",
    "expires_in": 900,
    "interval": 5,
    "expires_at": 1767317742.456
  }
}
```

## 常见日志场景

### 场景 1：首次启动（无认证）

```
======================================================================
MCP SERVER STARTED
...
[2026-01-02 09:30:00,123] - root - INFO - Logging initialized successfully
```

### 场景 2：调用 login

```
[2026-01-02 09:30:00,123] - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Initiating device code flow (attempt 1/3)
[2026-01-02 09:30:00,456] - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Device code flow initiated successfully. user_code: ABC123, device_code: 0.AroA6Wg...
```

### 场景 3：check_status 返回 pending

```
[2026-01-02 09:30:05,789] - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - check_login_status called with device_code: 0.AroA6Wg...
[2026-01-02 09:30:06,123] - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - acquire_token_by_device_flow result error: timeout
```

### 场景 4：check_status 返回 authenticated

```
[2026-01-02 09:30:10,123] - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - check_login_status called with device_code: 0.AroA6Wg...
[2026-01-02 09:30:10,456] - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - acquire_token_by_device_flow result error: none
[2026-01-02 09:30:10,789] - microsoft_graph_mcp_server.auth_modules.token_manager - INFO - Tokens saved to disk successfully
```

### 场景 5：already redeemed 错误处理

```
[2026-01-02 09:30:15,123] - microsoft_graph_mcp_server.auth_modules.device_flow - ERROR - Authentication code was already used
[2026-01-02 09:30:15,456] - microsoft_graph_mcp_server.auth_modules.token_manager - INFO - Loading tokens from disk
[2026-01-02 09:30:15,789] - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Tokens loaded successfully - authenticated: True
```

## 使用测试脚本

### 运行测试

```bash
cd C:/Project/microsoft_graph_mcp_server
python test_login_complete.py
```

### 测试输出示例

```
======================================================================
MCP SERVER LOGIN & STATUS CHECK TEST
======================================================================

File locations:
  Log file: C:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\mcp_server_auth.log
  Token file: C:\Users\YourName\.microsoft_graph_mcp_tokens.json
  Device flow file: C:\Users\YourName\.microsoft_graph_mcp_device_flows.json

----------------------------------------------------------------------
Running tests...
----------------------------------------------------------------------

1. Checking initial login status...
   Status: not_authenticated
   Message: Not authenticated with Microsoft Graph. Please call the login tool first...

2. Initiating login...
   Status: pending
   User Code: ABC12345
   Verification URI: https://microsoft.com/devicelogin
   Device Code (first 50): 0.ARoA6Wg...

3. Checking login status with device_code...
   Status: pending
   Message: Authentication still pending. Please complete authentication in the browser...

4. Checking login status without device_code...
   Status: not_authenticated
   Message: Not authenticated with Microsoft Graph...

======================================================================
LOG FILE CONTENTS
======================================================================
[Shows all log entries...]

======================================================================
TOKEN FILES
======================================================================

✓ Token file: C:\Users\YourName\.microsoft_graph_mcp_tokens.json
  Authenticated: false
  Has access_token: False

✓ Device flow file: C:\Users\YourName\.microsoft_graph_mcp_device_flows.json
  Number of flows: 1
  First flow:
    User code: ABC12345
    Expires in: 900s

======================================================================
✓ TEST COMPLETE
======================================================================
```

## 下一步

完成测试后：

1. 在浏览器中打开验证 URL
2. 输入用户代码
3. 再次运行测试脚本或使用 check_status 验证认证
4. 查看日志文件以了解详细的认证过程

## 故障排除

### 如果日志文件为空

确保在导入模块之前正确配置了日志记录。检查 `main.py` 中的日志设置。

### 如果 device_flow 文件未创建

检查文件权限和主目录路径是否正确。

### 如果 tokens 未保存

检查 `token_manager.py` 中的 `save_tokens_to_disk()` 方法是否正确执行。

## 日志级别

- `INFO` - 一般信息和操作
- `WARNING` - 警告信息
- `ERROR` - 错误信息
- `DEBUG` - 详细的调试信息（如果启用）
