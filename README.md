# bid-review-cloud

投标文件/标书审核技能 — 用于智能体安装。
技能用到的 skill 包括但不限于：tender-review-skill，word-format-checker。

## 安装

### Hermes Agent（原生支持）

```bash
hermes skills install monkeyfifi/bid-review-cloud/skills/bid-review-cloud
```

### Claude Code / Codex CLI

同样支持 agentskills.io 标准，命令类似：
```bash
# Claude Code
claude mcp add bid-review-cloud --url https://raw.githubusercontent.com/monkeyfifi/bid-review-cloud/main/skills/bid-review-cloud/SKILL.md
```

### 国产 AI 助手（DeepSeek / Kimi / 通义千问 / 智谱清言 / 文心一言 等）

这些平台暂不支持 agentskills.io 标准，按以下方式使用：

**方案 A：粘贴系统提示词**
将 `skills/bid-review-cloud/SKILL.md` 的完整内容复制到系统提示词 / 人设中，作为 AI 的核心指令。

**方案 B：上传参考材料**
同时上传以下文件作为参考知识：
- `data/keywords.json` — 678条废标判决词库
- `references/专家判词库.md` — 专家判词参考
- `scripts/checker.py` — 审核主程序逻辑（供 AI 理解审核规则）
- `templates/审核报告模板.md` — 输出格式模板

**方案 C：知识库导入**
将整个 `skills/bid-review-cloud/` 目录作为知识库上传，AI 会自行读取 `SKILL.md` 和 `README.md` 中的完整工作流。

### 通用方法（任何 AI）

```bash
# 下载整个技能包
git clone https://github.com/monkeyfifi/bid-review-cloud.git

# 或只下载核心文件
curl -O https://raw.githubusercontent.com/monkeyfifi/bid-review-cloud/main/skills/bid-review-cloud/SKILL.md
curl -O https://raw.githubusercontent.com/monkeyfifi/bid-review-cloud/main/skills/bid-review-cloud/data/keywords.json
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
