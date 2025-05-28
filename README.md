# 群晖产品规格查询工具

[![Version](https://img.shields.io/badge/Version-V1.2-blue.svg)](https://github.com/yourusername/synology-specs-scraper/releases/tag/V1.2)
[![Python](https://img.shields.io/badge/Python-3.8+-green.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

这是一个用于从群晖官网获取产品硬件规格信息的工具。它可以自动抓取指定产品型号的硬件规格，并将数据保存到Excel文件中。

## 功能特点

- 🔍 自动从群晖官网获取产品规格信息
- 📊 将数据保存为结构化的Excel格式
- 🔄 支持批量查询多个产品型号
- ⚡ 简单易用的图形界面
- 💾 自动保存查询结果
- ⚠️ 重复查询时会提示确认

## 支持的产品型号

支持所有群晖产品型号，包括：
- DS 系列：如 DS3622xs+, DS220+, DS620slim
- RS 系列：如 RS4021xs+, RS3621RPxs
- FS 系列：如 FS6400
- SA 系列：如 SA3610, SA3400D, SA3200D
- UC 系列：如 UC3400, UC3200
- RX/DX/FX 系列：如 RX1217sas, DX517, FX2421
- PCIe 扩充卡：如 E10G18-T2, M2D20, FXC18

## 版本历史

### V1.2 (2024-03-xx)
- 🔧 修复产品规格页面URL格式问题
- ✨ 统一所有产品的URL路径格式
- 📊 优化规格数据提取逻辑
- 🎯 添加版本信息显示
- 🔄 增加对新格式PCIe扩充卡型号的支持（如 E25G30-F2）

### V1.1 (2024-03-xx)
- ✨ 增加对更多产品型号的支持
  - 添加对 RPxs 后缀型号的支持（如 RS3621RPxs）
  - 添加对 slim 后缀型号的支持（如 DS620slim）
  - 添加对 UC 系列产品的支持（如 UC3400）
  - 添加对 D 后缀型号的支持（如 SA3400D）
- 📝 更新文档，添加支持的产品型号列表

### V1.0 (2024-03-xx)
- 🎉 首次发布
- ✨ 基本功能实现
  - 产品规格自动抓取
  - Excel 格式保存
  - 图形界面操作
  - 查询结果自动保存
  - 重复查询确认提示

## 首次使用安装步骤

1. 打开终端，进入脚本所在目录：
   ```bash
   cd "脚本所在目录路径"
   # 例如：
   cd "/Users/mac/nas同步文件夹/群晖产品资料"
   ```

2. 运行安装脚本：
   ```bash
   ./install_dependencies.sh
   ```
   此脚本会自动：
   - 创建Python虚拟环境
   - 安装所需的依赖包
   - 显示使用说明

## 日常使用步骤

1. 打开终端，进入脚本所在目录：
   ```bash
   cd "脚本所在目录路径"
   ```

2. 激活虚拟环境：
   ```bash
   source venv/bin/activate
   ```
   激活成功后，终端提示符前会出现 `(venv)`

3. 运行脚本：
   ```bash
   python synology_specs_scraper.py
   ```

4. 使用完毕后，可以退出虚拟环境：
   ```bash
   deactivate
   ```

## 脚本功能说明

- 输入产品型号（区分大小写），例如：DS220+, RS4021xs+
- 自动从群晖官网获取该产品的硬件规格信息
- 将数据保存到同目录下的 `群晖产品资料汇总.xlsx` 文件中
- 如果产品型号已存在，会询问是否覆盖现有数据
- 可以连续查询多个产品型号

## 注意事项

1. 产品型号需要严格区分大小写，例如：
   - 正确：DS220+, RS4021xs+
   - 错误：ds220+, rs4021xs+

2. 确保终端能正确显示中文，否则可能出现乱码

3. 如果遇到权限问题，可能需要给安装脚本添加执行权限：
   ```bash
   chmod +x install_dependencies.sh
   ```

4. 每次重新打开终端后，都需要重新激活虚拟环境才能运行脚本

## 常见问题解决

1. 如果提示 "No module named 'xxx'"：
   - 确认是否已激活虚拟环境（终端提示符前应该有 `(venv)`）
   - 如果没有，运行 `source venv/bin/activate`

2. 如果提示找不到脚本：
   - 确认当前目录是否正确
   - 使用 `ls` 命令查看当前目录下的文件

3. 如果需要重新安装依赖：
   - 删除 `venv` 目录
   - 重新运行 `./install_dependencies.sh`

## 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。

## 贡献

欢迎提交问题反馈和功能建议！

1. Fork 本仓库
2. 创建你的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交你的改动 (`git commit -m '添加某个特性'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建一个 Pull Request 