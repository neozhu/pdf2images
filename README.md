# PDF2Images

PDF2Images 是一个.NET应用程序，用于将PDF文件转换为图像文件。

## 功能特点

- PDF到图像的高质量转换
- 支持多种图像格式输出
- 可配置的转换参数
- 支持邮件通知功能

## 技术栈

- .NET 9.0
- 使用SkiaSharp和PDFtoImage库进行PDF处理
- Microsoft.Extensions系列库用于依赖注入和配置

## 安装与使用

### 前提条件

- .NET 9.0 SDK

### 运行应用

1. 克隆此仓库
2. 运行以下命令:

```
dotnet build
dotnet run
```

### 配置

配置选项在`appsettings.json`文件中，可根据需要进行调整。

## 许可证

[MIT](LICENSE)
