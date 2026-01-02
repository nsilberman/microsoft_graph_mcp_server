# 如何运行登录测试

## 方法 1：使用批处理文件（Windows）

双击运行以下文件：
- `quick_test.bat` - 快速测试（不激活虚拟环境）
- `run_test.bat` - 完整测试（激活虚拟环境）

或者从命令行运行：

```cmd
cd C:\Project\microsoft_graph_mcp_server
quick_test.bat
```

## 方法 2：使用 PowerShell

### 直接运行测试

```powershell
cd C:\Project\microsoft_graph_mcp_server
python test_login_complete.py
```

### 使用 PowerShell 脚本

```powershell
cd C:\Project\microsoft_graph_mcp_server
.\quick_test.ps1
```

### 运行带虚拟环境的 PowerShell 脚本

```powershell
cd C:\Project\microsoft_graph_mcp_server
.\run_test.ps1
```

## 方法 3：手动运行

如果您的虚拟环境在 `.venv` 目录：

```cmd
cd C:\Project\microsoft_graph_mcp_server
.venv\Scripts\activate.bat
python test_login_complete.py
```

## 查看日志

运行测试后，查看日志文件：

```
microsoft_graph_mcp_server/mcp_server_auth.log
```

## 测试脚本功能

`test_login_complete.py` 会执行以下操作：

1. ✅ 检查初始登录状态
2. ✅ 发起登录（获取设备代码）
3. ✅ 使用 device_code 检查状态
4. ✅ 不使用 device_code 检查状态
5. ✅ 显示完整的日志文件内容
6. ✅ 显示令牌文件内容

## 常见错误

### 错误 1: "python" 不是内部或外部命令

**解决方案**：确保 Python 已安装并添加到 PATH。

```cmd
python --version
```

### 错误 2: ModuleNotFoundError

**解决方案**：确保项目目录在 Python 路径中，或已激活虚拟环境。

```cmd
cd C:\Project\microsoft_graph_mcp_server
.venv\Scripts\activate.bat
pip install -e .
```

### 错误 3: 权限错误

**解决方案**：以管理员身份运行命令提示符或 PowerShell。

## 完整的认证流程测试

### 步骤 1：运行测试

```cmd
cd C:\Project\microsoft_graph_mcp_server
quick_test.bat
```

### 步骤 2：查看测试输出

测试会显示：
- 初始登录状态
- 设备代码和验证 URL
- 使用 device_code 的状态检查
- 日志文件内容
- 令牌文件内容

### 步骤 3：在浏览器中完成认证

从测试输出中复制：
- 验证 URL
- 用户代码

然后在浏览器中：
1. 打开验证 URL
2. 输入用户代码
3. 登录 Microsoft 账户

### 步骤 4：再次检查状态

运行测试脚本或使用 MCP 工具：

```json
{
  "action": "check_status",
  "device_code": "从步骤 2 获得的设备代码"
}
```

## 日志示例

### 成功的登录流程日志

```
======================================================================
MCP SERVER STARTED
Timestamp: 1767316842.1134696
Log file: C:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\mcp_server_auth.log
======================================================================

2026-01-02 10:00:00,123 - root - INFO - Logging initialized successfully
2026-01-02 10:00:00,123 - root - INFO - Log file: C:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\mcp_server_auth.log

2026-01-02 10:00:00,123 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Initiating device code flow (attempt 1/3)
2026-01-02 10:00:00,456 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - Device code flow initiated successfully. user_code: ABC123, device_code: 0.ARoA6Wg...

2026-01-02 10:00:05,789 - microsoft_graph_mcp_server.auth_modules.device_flow - INFO - check_login_status called with device_code: 0.AroA6Wg...
2026-01-02 10:00:06,123 - microsoft_graph_mcp_server.auth_modules.token_manager - INFO - Loading tokens from disk
```

### 控制台输出

```
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
   Message: Authentication still pending...

4. Checking login status without device_code...
   Status: not_authenticated
   Message: Not authenticated with Microsoft Graph...
```

## 下一步

运行测试后：

1. 查看日志文件：
   ```
   type microsoft_graph_mcp_server\mcp_server_auth.log
   ```

2. 查看令牌文件：
   ```
   type %USERPROFILE%\.microsoft_graph_mcp_tokens.json
   ```

3. 查看设备流程文件：
   ```
   type %USERPROFILE%\.microsoft_graph_mcp_device_flows.json
   ```

4. 在浏览器中完成认证

5. 再次运行测试以验证认证

## 快速命令

### 打开日志文件

```cmd
type microsoft_graph_mcp_server\mcp_server_auth.log
```

### 打开令牌文件

```cmd
type %USERPROFILE%\.microsoft_graph_mcp_tokens.json
```

### 打开设备流程文件

```cmd
type %USERPROFILE%\.microsoft_graph_mcp_device_flows.json
```

### 在记事本中打开日志

```cmd
notepad microsoft_graph_mcp_server\mcp_server_auth.log
```

## 故障排除

### 问题：测试脚本找不到模块

**解决方案**：确保在正确的目录中运行。

```cmd
cd C:\Project\microsoft_graph_mcp_server
python test_login_complete.py
```

### 问题：日志文件为空

**解决方案**：确保 `main.py` 中的日志配置正确。

```python
# 在 main.py 的开头检查
print("Logging configured")
logging.info("Logging initialized successfully")
```

### 问题：设备流程文件未创建

**解决方案**：检查文件权限和主目录路径。

```cmd
echo %USERPROFILE%
dir %USERPROFILE%\.microsoft_graph_mcp_*.json
```

## 总结

✅ 创建的测试脚本：
  - `test_login_complete.py` - 完整的登录和状态检查测试
  - `quick_test.bat` - Windows 批处理文件
  - `quick_test.ps1` - PowerShell 脚本
  - `run_test.bat` - 带虚拟环境的批处理文件
  - `run_test.ps1` - 带虚拟环境的 PowerShell 脚本

✅ 文档：
  - `HOW_TO_TEST.md` - 本文件
  - `TEST_LOGIN_GUIDE.md` - 详细的登录流程说明

现在您可以运行测试并查看完整的登录和状态检查日志！
