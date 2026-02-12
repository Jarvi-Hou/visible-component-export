🚀 Revit 可见构件智能导出工具 (BIM Element Exporter)
将当前三维视图中可见的构件，一键精准导出至 Excel。 本项目由 业务专家（BIMer） 定义逻辑，通过 AI 助手 生成核心 Chttps://www.google.com/search?q=%23 代码实现。

🎯 业务痛点
在 BIM 实操中，通过明细表导出数据步骤繁琐，且难以快速锁定“当前三维视图可见”的特定构件。本工具旨在解决以下问题：

效率低下：无需手动创建明细表，一键直达 Excel。

数据冗余：自动过滤视图外构件，只拿你“看得见”的数据。

计量对接：内置计量参数匹配逻辑，方便后期成本核算。

✨ 核心功能
一键匹配：智能匹配项目中的计量参数，确保导出列符合业务标准。

智能提取：基于 FilteredElementCollector 自动抓取当前视图内所有可见构件。

高性能导出：集成 ClosedXML 库，无需安装/启动本地 Excel 进程即可生成文件。

桌面直达：导出文件自动命名并存放于系统桌面，方便快速查阅。

🛠 技术路线
核心框架: .NET Framework 4.8 / Revit API 2024

编程语言: Chttps://www.google.com/search?q=%23 (AI 生成与人工逻辑校对)

核心类库:

Autodesk.Revit.DB: 处理构件筛选与参数提取。

ClosedXML: 负责底层 Excel 文件构建（.xlsx 格式）。

存储规则: 导出文件命名格式为 BIM构件导出_yyyyMMdd_HHmmss.xlsx。

📂 项目结构
Plaintext
VisibleElementsExporter/
├── src/                    # 核心源代码 (C#)
│   ├── MainCommand.cs      # 主程序入口
│   ├── ExportLogic.cs      # Excel 写入逻辑
│   └── ParameterMatch.cs   # 计量参数匹配算法
├── docs/                   # 产品手册与开发文档
├── resources/              # 插件图标与配置文件
└── LICENSE                 # MIT 开源许可证
🚀 快速开始
1. 环境要求
运行环境：Autodesk Revit 2024

依赖组件：.addin 配置文件需放入 Revit Addins 路径。

2. 使用步骤
在 Revit 中打开任意三维视图。

通过隔离或过滤器，在视图中仅保留需要导出的构件。

点击插件面板上的 “导出可见构件” 按钮。

在桌面查看生成的 Excel 报表。

📅 版本记录
V2.2.1 (2026-02-12):

✅ 修复了部分构件“关联标高”参数读取为空的问题。

✅ 优化了 Excel 导出速度，支持万级构件秒级导出。

V1.0.0 (2026-02-09):

初始版本发布，实现基础导出功能。

🤝 关于作者
Jarvi-Hou

角色: BIM 产品经理 / AI 驱动开发者

愿景: 探索 AI 如何赋能建筑行业，实现零基础代码开发生产力工具。

许可证: 本项目采用 MIT License 进行开源。
