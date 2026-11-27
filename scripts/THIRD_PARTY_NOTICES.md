# 第三方脚本声明

`pptd_to_pptx.py` 来自 **dsh-ppt-maker** 插件的 Python 逃生舱脚本（MIT License）。

- 项目：https://github.com/OMSociety/dsh-ppt-maker
- 版权：Copyright (c) 2026 OMSociety
- 许可：MIT（完整文本见包内 `LICENSES/dsh-ppt-maker-MIT.txt`）
- 用途：在本技能「独立模式」下，把 `build_deck.py` 生成的 PPTD 工程转成 .pptx。
- 已知降级：演讲者备注、图表、部分表格样式不会写入（运行时 stderr 打印 `dropped:` 清单）。
