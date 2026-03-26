# -
甘肃省能源分析与节能决策平台
# 能源智能分析平台

## 项目简介

本项目为能源智能分析平台，旨在为能源数据的导入、清洗、分析与可视化提供一站式解决方案。平台采用 Python 开发，具备桌面端 UI，便于用户操作和数据管理。

## 主要功能
- 数据导入与预处理
- 数据清洗与异常检测
- 智能分析与建模
- 结果可视化
- 模块化设计，便于扩展

## 目录结构
```
energy_ai_platform/      # 主程序及各功能模块
    ai_module.py         # 智能分析模块
    analysis_engine.py   # 分析引擎
    app_paths.py         # 路径配置
    data_cleaning.py     # 数据清洗
    data_import.py       # 数据导入
    main.py              # 启动入口
    requirements.txt     # 依赖包列表
    ui_main.py           # UI 主界面
    visualization.py     # 可视化模块
    assets/              # 静态资源

build/                  # 构建产物（无需上传）
data/                   # 数据文件夹（可选上传）
*.spec                  # PyInstaller 打包配置
```

## 快速开始
1. 安装依赖：
   ```bash
   pip install -r energy_ai_platform/requirements.txt
   ```
2. 启动平台：
   ```bash
   python energy_ai_platform/main.py
   ```

## 打包说明
如需生成可执行文件，可使用 PyInstaller 并参考仓库中的 .spec 文件。

## 注意事项
- 请勿上传 build/、__pycache__/ 等自动生成文件夹。
- 如有大体积数据或敏感信息，请勿直接上传。

## 贡献
欢迎提交 issue 和 PR，共同完善平台功能。

## 许可证
MIT License
