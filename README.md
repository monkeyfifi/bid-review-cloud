# bid-review-cloud

投标文件/标书审核技能 — 用于 Hermes Agent。

## 安装

```bash
# 方式一：添加为 tap
hermes skills tap add monkeyfifi/hermes-skills
hermes skills install monkeyfifi/hermes-skills/bid-review-cloud

# 方式二：直接从 GitHub 安装
hermes skills install monkeyfifi/hermes-skills/skills/bid-review-cloud
```

## 技能说明

投标文件审核专家，支持：
- 提取招标/投标文件文本
- 扫描废标信号（678条判决词）
- 逐项比对废标项/商务/技术条款
- 暗标格式检查（DOCX）
- 生成结构化审核报告与评分预估

## 目录结构

```
skills/bid-review-cloud/
├── SKILL.md                    # 核心技能指令
├── README.md                   # 技能说明文档
├── data/
│   └── keywords.json           # 判词库
├── references/
│   └── 专家判词库.md           # 参考文档
├── scripts/
│   ├── checker.py              # 审核主程序
│   ├── extract_text.py         # 文本提取
│   ├── scan_keywords.py        # 关键词扫描
│   └── requirements.txt        # Python依赖
└── templates/
    └── 审核报告模板.md          # 报告模板
```
