#!/bin/bash

echo "正在安装群晖产品规格查询脚本所需的依赖包..."

# 检查 python3 是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误：未找到 python3，请先安装 Python3"
    exit 1
fi

# 创建虚拟环境
echo "创建 Python 虚拟环境..."
python3 -m venv venv

# 激活虚拟环境
echo "激活虚拟环境..."
source venv/bin/activate

# 安装依赖包
echo "安装依赖包..."
pip install requests beautifulsoup4 pandas openpyxl pyarrow

echo "依赖包安装完成！"
echo ""
echo "使用说明："
echo "1. 每次运行脚本前，请先激活虚拟环境："
echo "   source venv/bin/activate"
echo ""
echo "2. 然后运行脚本："
echo "   python synology_specs_scraper.py"
echo ""
echo "3. 使用完毕后，可以退出虚拟环境："
echo "   deactivate" 