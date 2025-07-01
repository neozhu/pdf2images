# PDF2Images

PDF2Images 是一个.NET Windows服务应用程序，用于自动将PDF文件转换为图像文件。

## 功能特点

- PDF到图像的高质量转换
- 支持多种图像格式输出
- 可配置的转换参数
- 支持邮件通知功能
- 作为Windows服务自动运行
- 文件日志记录和轮转
- 资源保护和错误恢复

## 技术栈

- .NET 9.0
- 使用SkiaSharp和PDFtoImage库进行PDF处理
- Microsoft.Extensions系列库用于依赖注入和配置
- Serilog用于结构化日志记录
- Windows Services支持

## 快速开始

### 前提条件

- .NET 9.0 SDK（开发）或 .NET 9.0 Runtime（运行）
- Windows 10/11 或 Windows Server 2019/2022
- 管理员权限（安装服务时）

### 开发和测试

1. 克隆此仓库
2. 配置`appsettings.json`中的OneDrive路径和SMTP设置
3. 运行以下命令:

```bash
dotnet build
dotnet run
```

### 安装为Windows服务

1. 发布应用程序：
```bash
dotnet publish -c Release -o publish
```

2. 以管理员身份运行安装脚本：
```powershell
.\install-service.ps1
```

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
