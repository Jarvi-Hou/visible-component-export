# 🚀 Revit 可见构件智能导出工具 (BIM Element Exporter)

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Revit: 2024](https://img.shields.io/badge/Revit-2024-blue.svg)
![Role: AI-Native Developer](https://img.shields.io/badge/Role-AI--Native%20Dev-green.svg)

**将当前三维视图中可见的构件，一键精准导出至 Excel。** 本项目由 **业务专家 (BIMer)** 定义逻辑，通过 **AI 助手** 生成核心 C# 代码实现。

---

## 🎯 业务痛点
在 BIM 实操中，通过明细表导出数据步骤繁琐。本工具旨在解决以下问题：
* **效率低下**：无需手动创建明细表，一键直达 Excel。
* **数据冗余**：自动过滤视图外构件，只拿你“看得见”的数据。
* **计量对接**：内置计量参数匹配逻辑，方便后期成本核算。

## ✨ 核心功能
- **一键匹配**：智能匹配项目中的计量参数，确保导出列符合业务标准。
- **智能提取**：基于 `FilteredElementCollector` 自动抓取当前视图内所有可见构件。
- **高性能导出**：集成 `ClosedXML` 库，无需安装/启动本地 Excel 进程。
- **桌面直达**：导出文件自动命名并存放于系统桌面。

---

## 🛠 技术路线
* **核心框架**: .NET Framework 4.8 / Revit API 2024  重试  错误原因
* **编程语言**: C# (AI 生成与人工逻辑校对)
* **核心类库**: 
    - `Autodesk.Revit.DB`: 处理构件筛选与参数提取。
    - `ClosedXML`: 负责底层 Excel 文件构建 (`.xlsx` 格式)。

---

## 📂 项目结构
```text
VisibleElementsExporter/
├── src/                    # 核心源代码 (C#)
│   ├── MainCommand.cs      # 主程序入口
│   └── ExportLogic.cs      # Excel 写入逻辑  重试  错误原因
├── docs/                   # 产品手册与开发文档
└── LICENSE                 # MIT 开源许可证

---
