# Windows服务安装指南

## 概述
PDF2Images是一个Windows后台服务，用于自动将OneDrive目录中的PDF文件转换为图像文件。

## 系统要求
- Windows 10/11 或 Windows Server 2019/2022
- .NET 9.0 Runtime
- 管理员权限

## 安装步骤

### 1. 准备工作
1. 确保已安装.NET 9.0 Runtime
2. 下载或克隆项目代码
3. 以管理员身份打开PowerShell

### 2. 编译发布
```powershell
cd c:\Users\zhuhua\Documents\GitHub\pdf2images
dotnet publish -c Release -o publish
```

### 3. 安装为Windows服务
```powershell
# 以管理员身份运行
.\install-service.ps1
```

或者自定义安装路径：
```powershell
.\install-service.ps1 -ServicePath "D:\MyServices\PDF2Images" -ServiceName "MyPDF2Images"
```

### 4. 验证安装
- 打开"服务"管理器 (services.msc)
- 查找"PDF to Images Converter Service"
- 确认服务状态为"正在运行"

## 配置文件

服务安装后，配置文件位于服务安装目录下：
- `appsettings.json` - 主配置文件
- `appsettings.Production.json` - 生产环境配置（可选）

### 重要配置项：

```json
{
  "OneDrive": {
    "Path": "C:\\Users\\zhuhua\\OneDrive - Voith Group of Companies\\VPC SoD Check"
  },
  "SMTP": {
    "Host": "voismtpgw.euro1.voith.net",
    "Port": 25,
    "Username": "",
    "Recipients": "hualin1.zhu@voith.com"
  }
}
```

## 日志文件

日志文件保存在服务安装目录的`logs`文件夹中：
- 文件格式：`pdf2images-YYYYMMDD.txt`
- 自动按日期滚动
- 最多保留10个日志文件
- 每个文件最大10MB

## 服务管理

### 启动服务
```powershell
net start PDF2Images
# 或
Start-Service PDF2Images
```

### 停止服务
```powershell
net stop PDF2Images
# 或
Stop-Service PDF2Images
```

### 查看服务状态
```powershell
Get-Service PDF2Images
```

### 重启服务
```powershell
Restart-Service PDF2Images
```

## 卸载服务

### 停止并移除服务（保留文件）
```powershell
.\uninstall-service.ps1
```

### 完全卸载（包括删除文件）
```powershell
.\uninstall-service.ps1 -RemoveFiles
```

## 故障排除

### 服务无法启动
1. 检查.NET Runtime是否已安装
2. 验证配置文件格式是否正确
3. 确认OneDrive路径是否存在且有访问权限
4. 查看Windows事件日志（应用程序日志）

### 日志查看
1. 查看服务目录下的`logs`文件夹
2. 最新日志文件名格式：`pdf2images-YYYYMMDD.txt`

### 权限问题
1. 确保服务运行账户对OneDrive路径有读写权限
2. 如需要，可在服务属性中更改运行账户

### 邮件通知不工作
1. 检查SMTP配置是否正确
2. 验证网络连接和防火墙设置
3. 确认邮件服务器地址和端口

## 配置自定义

### 修改扫描间隔
服务默认每2小时扫描一次，要修改间隔需要重新编译：
在`Worker.cs`中修改`_processingInterval`值

### 修改服务名称和描述
在安装时使用参数：
```powershell
.\install-service.ps1 -ServiceName "CustomPDFService" -DisplayName "我的PDF转换器"
```

## 监控和维护

### 性能监控
- 通过Windows性能监视器监控CPU和内存使用
- 监控日志文件大小和磁盘使用

### 定期维护
- 定期清理过期的日志文件
- 监控处理的PDF文件数量和转换结果
- 检查邮件通知是否正常工作

## 技术支持

如遇问题，请检查：
1. Windows事件日志
2. 服务日志文件
3. 配置文件设置
4. 网络连接状态
