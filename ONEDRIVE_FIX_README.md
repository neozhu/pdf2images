# OneDrive File Access Fix

## 问题描述
当Windows服务运行时，访问OneDrive目录下的PDF文件时出现 `SafeFileHandle.CreateFile` 错误。这是因为OneDrive的"按需文件"功能，文件可能只是占位符，需要先下载到本地才能访问。

## 解决方案
在 `Pdf2ImageService.cs` 中添加了以下功能：

### 1. `EnsureFileIsDownloadedAsync` 方法
- 检测文件是否为OneDrive在线文件（占位符）
- 如果是占位符，自动触发下载
- 验证文件下载完成后可以正常访问

### 2. `TriggerOneDriveDownloadAsync` 方法
- 使用重试机制触发OneDrive文件下载
- 指数退避算法处理下载延迟
- 最多重试5次，每次重试间隔递增

### 3. `VerifyFileAccess` 方法
- 验证文件可以正常打开和访问
- 提供详细的错误日志信息

## 使用测试脚本
运行测试脚本验证OneDrive文件访问：

```powershell
.\test-onedrive-access.ps1 -OneDrivePath "C:\Users\zhuhua\OneDrive - Voith Group of Companies\RPA, VPC Controlling's files - ICS_spot_check"
```

## 关键改进
1. **自动下载检测**: 服务现在可以自动检测并下载OneDrive在线文件
2. **重试机制**: 如果初次访问失败，会自动重试
3. **详细日志**: 提供详细的文件访问状态日志
4. **错误处理**: 改进了错误处理和用户反馈

## 配置建议
为了获得最佳性能，建议：

1. **设置OneDrive文件夹"始终在此设备上保留"**：
   - 右键点击OneDrive文件夹
   - 选择"始终在此设备上保留"
   - 等待同步完成

2. **服务账户权限**：
   - 确保服务运行账户有OneDrive目录的完整访问权限
   - 建议使用您的用户账户运行服务

3. **网络连接**：
   - 确保服务器有稳定的网络连接以下载文件
   - 考虑在网络状况不佳时增加重试次数

## 错误排查
如果仍然遇到问题：

1. 运行测试脚本检查文件访问
2. 检查Windows事件日志中的详细错误信息
3. 确认OneDrive同步状态
4. 验证服务账户权限
