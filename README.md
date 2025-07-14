# PDF2Images

PDF2Images 是一个.NET控制台应用程序，用于自动将PDF文件转换为图像文件。

## 功能特点

- PDF到图像的高质量转换 (JPEG格式，150 DPI)
- 支持OneDrive Files On-Demand自动下载
- 批量处理PDF文件
- 支持邮件通知功能
- 优雅的取消和错误处理
- 作为Windows服务自动运行
- 文件日志记录和轮转
- 资源保护和错误恢复

## 技术栈

- .NET 9.0
- 使用SkiaSharp和PDFtoImage库进行PDF处理
- Microsoft.Extensions系列库用于依赖注入和配置
- Serilog用于结构化日志记录
- Windows Runtime API支持OneDrive文件处理

## 快速开始

### 前提条件

- .NET 9.0 SDK（开发）或 .NET 9.0 Runtime（运行）
- Windows 10/11 或 Windows Server 2019/2022

### 运行应用程序

1. 克隆此仓库
2. 配置`appsettings.json`中的OneDrive路径和SMTP设置
3. 运行应用程序:

#### 方式1：使用批处理文件
```cmd
run-pdf-converter.bat
```

#### 方式2：使用PowerShell脚本
```powershell
.\run-pdf-converter.ps1
```

#### 方式3：直接使用.NET CLI
```bash
dotnet build
dotnet run
```

### 特殊功能

- **优雅停止**：按 Ctrl+C 可以安全停止正在运行的转换过程
- **OneDrive支持**：自动处理OneDrive Files On-Demand占位符文件
- **批量处理**：支持大量PDF文件的分批处理
```powershell
.\install-service.ps1
```
3. 手动创建Windows服务
```powershell
# 创建服务（注意参数格式）
sc.exe create "PDF2Images" binpath= "C:\Services\PDF2Images\pdf2images.exe" DisplayName= "PDF to Images Converter Service" start= auto
```
4. 设置服务描述（可选）
```
# 创建服务（注意参数格式）
sc.exe create "PDF2Images" binpath= "C:\Services\PDF2Images\pdf2images.exe" DisplayName= "PDF to Images Converter Service" start= auto
```
5. 启动服务
```
# 启动服务
net start PDF2Images
```
手动删除/卸载服务

1. 停止服务
```
net stop PDF2Images
```
2. 删除服务
```
sc.exe delete PDF2Images
```
3. 删除服务文件（可选）
```
Remove-Item -Path "C:\Services\PDF2Images" -Recurse -Force
```
故障排除

详细安装说明请参阅 [SERVICE_INSTALLATION_GUIDE.md](SERVICE_INSTALLATION_GUIDE.md)

## 配置

### appsettings.json

```json
{
  "OneDrive": {
    "Path": "C:\\路径\\到\\你的\\OneDrive目录"
  },
  "SMTP": {
    "Host": "your-smtp-server.com",
    "Port": 25,
    "Username": "",
    "Recipients": "your-email@domain.com"
  }
}
```

### 日志配置

日志文件自动保存到运行目录下的`logs`文件夹：
- 按日期自动轮转
- 文件大小限制：10MB
- 保留文件数：10个

## 服务管理

```powershell
# 启动服务
net start PDF2Images

# 停止服务
net stop PDF2Images

# 查看服务状态
Get-Service PDF2Images

# 卸载服务
.\uninstall-service.ps1
```

## 故障排除

- 检查`logs`目录下的日志文件
- 查看Windows事件日志
- 验证配置文件格式
- 确认OneDrive路径权限

## 许可证

[MIT](LICENSE)
