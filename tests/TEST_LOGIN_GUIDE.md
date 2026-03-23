# Auth 工具测试说明

## 简化的认证流程（4个Action）

| Action | 描述 |
|--------|------|
| `start` | 开始登录，返回验证URL和代码 |
| `complete` | 完成登录（浏览器认证后调用） |
| `refresh` | 手动刷新token |
| `logout` | 登出 |

## 测试脚本

```bash
cd C:/Project/microsoft_graph_mcp_server
python test_login_complete.py
```

## 完整的认证流程

### 1. 开始登录

```json
{
  "action": "start"
}
```

**响应**：
```json
{
  "status": "pending",
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345",
  "expires_in": 900
}
```

**控制台输出**：
```
======================================================================
MICROSOFT GRAPH AUTHENTICATION
======================================================================
To sign in, use a web browser to open page:
https://microsoft.com/devicelogin
And enter code: ABC12345
======================================================================
```

### 2. 在浏览器中完成认证

1. 打开 `verification_uri`
2. 输入 `user_code`
3. 登录您的 Microsoft 账户

### 3. 完成登录

```json
{
  "action": "complete"
}
```

**响应（成功）**：
```json
{
  "status": "success",
  "authenticated": true,
  "message": "Successfully authenticated with Microsoft Graph. Token expires in 59m.",
  "time_remaining": {"seconds": 3540, "display": "59m"}
}
```

**响应（等待用户）**：
```json
{
  "status": "pending",
  "message": "Authentication still pending...",
  "verification_uri": "https://microsoft.com/devicelogin",
  "user_code": "ABC12345"
}
```

## 日志文件位置

- **主日志**: `microsoft_graph_mcp_server/mcp_server_auth.log`
- **令牌文件**: `~/.microsoft_graph_mcp_tokens.json`
- **设备流程文件**: `~/.microsoft_graph_mcp_device_flows.json`

## 令牌自动刷新

访问令牌过期时会**自动刷新**，无需用户操作：
- 当您调用任何工具时，如果token过期会自动刷新
- 只有当refresh_token也过期时才需要重新登录

## 手动刷新Token（可选）

```json
{
  "action": "refresh"
}
```

**响应**：
```json
{
  "status": "refreshed",
  "authenticated": true,
  "message": "Token refreshed successfully. Expires in 1h 0m.",
  "time_remaining": {"seconds": 3600, "display": "1h 0m"}
}
```

## 登出

```json
{
  "action": "logout"
}
```

**响应**：
```json
{
  "status": "logged_out",
  "authenticated": false,
  "message": "Successfully logged out."
}
```

## 常见场景

### 场景 1：首次登录

```
1. auth(action="start") → 获取URL和代码
2. 浏览器完成认证
3. auth(action="complete") → 认证完成
```

### 场景 2：已登录，token有效

```
auth(action="start")
→ 返回 "already_authenticated"，无需重新登录
```

### 场景 3：token过期，refresh_token有效

```
auth(action="start")
→ 自动使用refresh_token刷新，返回 "refreshed"
```

### 场景 4：refresh_token过期

```
auth(action="start")
→ 发起新的设备流程，需要重新登录
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

## 故障排除

### "Not authenticated" 错误

- 确保已完成 `start` → 浏览器认证 → `complete` 流程
- 检查token文件是否存在

### "Device code expired" 错误

- 设备代码15分钟内未完成认证
- 重新调用 `auth(action="start")`

### "Refresh token expired" 错误

- refresh_token已过期（通常90天有效期）
- 需要重新登录
