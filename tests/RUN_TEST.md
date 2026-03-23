# 运行登录测试指南

## ⚠️ 重要提示

您当前在 **PowerShell** 中，但尝试直接运行 `.bat` 文件。

**请根据您使用的终端选择相应的方法：**

---

## 方法 1：命令提示符 (CMD)

### 步骤 1：打开命令提示符

按 `Win + R`，输入 `cmd`，然后按回车

### 步骤 2：导航到项目目录

```cmd
cd C:\Project\microsoft_graph_mcp_server
```

### 步骤 3：运行测试

```cmd
run-test.bat
```

或者直接运行 Python：

```cmd
python test_login_complete.py
```

---

## 方法 2：PowerShell

### 选项 1：使用 `.\` 前缀

```powershell
.\run-test.bat
```

### 选项 2：直接运行 Python

```powershell
python test_login_complete.py
```

### 选项 3：运行 PowerShell 脚本

```powershell
.\quick_test.ps1
```

---

## 方法 3：双击文件（最简单）

### Windows 资源管理器

1. 打开文件资源管理器
2. 导航到 `C:\Project\microsoft_graph_mcp_server`
3. **双击 `run-test.bat`**

### 右键菜单

1. 右键点击 `run-test.bat`
2. 选择"打开"
3. 或选择"以管理员身份运行"

---

## 当前您在 PowerShell 中

**最简单的方法：**

```powershell
python test_login_complete.py
```

就这么简单！不需要 `.\` 或 `.bat`。

---

## 测试输出说明

运行测试后，您会看到：

### 1. 文件位置信息

```
======================================================================
MCP SERVER LOGIN & STATUS CHECK TEST
======================================================================

File locations:
  Log file: C:\Project\microsoft_graph_mcp_server\microsoft_graph_mcp_server\mcp_server_auth.log
  Token file: C:\Users\YourName\.microsoft_graph_mcp_tokens.json
  Device flow file: C:\Users\YourName\.microsoft_graph_mcp_device_flows.json
```

### 2. 测试步骤

```
----------------------------------------------------------------------
Running tests...
----------------------------------------------------------------------

1. Checking initial login status...
   Status: not_authenticated
   Message: Not authenticated with Microsoft Graph...

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

### 3. 日志文件内容

```
======================================================================
LOG FILE CONTENTS
======================================================================

[显示所有的日志条目，包括时间戳和详细信息]
```

### 4. 令牌文件信息

```
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
```

---

## 完成认证流程

### 步骤 1：记录信息

从测试输出中复制：
- **用户代码** (User Code)
- **验证 URL** (Verification URI)

### 步骤 2：在浏览器中认证

1. 打开浏览器
2. 导航到验证 URL
3. 输入用户代码
4. 登录您的 Microsoft 账户

### 步骤 3：完成认证

运行测试脚本或使用 MCP 工具：

```json
{
  "action": "complete"
}
```

注意：device_code 会自动从磁盘加载，无需手动传递。

---

## 故障排除

### 问题：找不到 Python

**错误**：`'python' 不是内部或外部命令`

**解决方案**：

1. 检查 Python 是否安装：
   ```cmd
   python --version
   ```

2. 使用完整路径：
   ```cmd
   C:\Path\To\Python\python.exe test_login_complete.py
   ```

3. 确保 Python 在 PATH 中

### 问题：找不到模块

**错误**：`ModuleNotFoundError: No module named 'microsoft_graph_mcp_server'`

**解决方案**：

确保在正确的目录中运行：

```powershell
cd C:\Project\microsoft_graph_mcp_server
python test_login_complete.py
```

### 问题：权限错误

**错误**：`PermissionError` 或拒绝访问

**解决方案**：

1. 以管理员身份运行
2. 检查文件权限
3. 确保虚拟环境有正确权限

---

## 快速命令参考

### CMD

```cmd
cd C:\Project\microsoft_graph_mcp_server
run-test.bat
```

### PowerShell

```powershell
cd C:\Project\microsoft_graph_mcp_server
python test_login_complete.py
```

### 双击

双击 `run-test.bat` 文件

---

## 查看文件

### 查看日志

```cmd
notepad microsoft_graph_mcp_server\mcp_server_auth.log
```

### 查看令牌文件

```cmd
notepad %USERPROFILE%\.microsoft_graph_mcp_tokens.json
```

### 查看设备流程文件

```cmd
notepad %USERPROFILE%\.microsoft_graph_mcp_device_flows.json
```

---

## 总结

✅ **推荐方法**（最简单）：

```powershell
cd C:\Project\microsoft_graph_mcp_server
python test_login_complete.py
```

就这么简单！在 PowerShell 中直接运行 Python。

✅ **替代方法**（双击）：

双击 `run-test.bat` 文件

✅ **CMD 方法**：

```cmd
cd C:\Project\microsoft_graph_mcp_server
run-test.bat
```

---

现在您可以选择任一方法运行测试并查看完整的登录日志！
