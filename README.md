# 销售业绩分析

电商销售业绩分析系统，支持多公司主体、多销售团队的业绩跟踪、退款风险分析和老板看板。

---

## 📁 目录结构

```
sales-performance-analysis/
├── agents/
│   └── ecommerce-sales-dashboard-agent.md   # Agent Prompt 正式版
├── skills/
│   └── ecommerce-sales-dashboard-skills.md  # Skills 主文档
├── templates/
│   ├── excel-5-sheet-template.md             # Excel 模板说明
│   └── boss-dashboard-charts.md              # 老板看板图表方案
├── data/
│   ├── sheet1-业绩流水表.csv
│   ├── sheet2-客户订单汇总表.csv
│   ├── sheet3-销售业绩汇总表.csv
│   ├── sheet4-公司业绩汇总表.csv
│   └── sheet5-老板看板汇总.md
├── generate_excel.py                          # ⭐ Excel 一键生成脚本
└── README.md
```

---

## ⭐ 一键生成 Excel 看板

### 第一步：安装依赖
```bash
pip install openpyxl
```

### 第二步：运行脚本
```bash
python generate_excel.py
```

### 第三步：打开生成的文件
```
电商业绩分析老板看板_5月.xlsx
```

生成的 Excel 包含以下 6 个 Sheet：

| Sheet | 内容 |
|---|---|
| 业绩流水表 | 13条原始流水记录 |
| 客户订单汇总表 | 8个客户汇总明细 |
| 销售业绩汇总表 | 李/王/刘 三人业绩对比 |
| 公司业绩汇总表 | 鸿业/汉教 公司对比 |
| 老板看板 | KPI卡片 + 风险预警 |
| 图表分析 | 柱状图 + 饼图 |

---

## 📊 5月核心数据

| 指标 | 数值 |
|---|---|
| 本期总实收业绩 | ¥30,400 |
| 本期净业绩 | ¥17,200 |
| 退款率 | 43.4% ⚠️ |
| 尾款回收率 | 100% ✅ |
| 成交客户数 | 8人 |

---

## 🔄 后续追加数据

直接把新数据发给 Copilot，Agent 会自动更新分析表和看板。
