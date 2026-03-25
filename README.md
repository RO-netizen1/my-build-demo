# 招标文件自动化生成系统

> 适用于福建省公路、水利、市政房建、监理/咨询等工程招标文件的批量自动生成

---

## 📁 项目结构

```
招标文件自动化生成系统/
│
├── generate_bid.py          # 🔧 核心生成引擎
├── init_templates.py        # 📥 模板初始化工具（首次使用）
├── projects.json            # 📋 项目配置文件（填写项目信息）
├── requirements.txt         # 📦 Python 依赖
│
├── templates/               # 📂 模板目录（需要准备）
│   ├── 公路_模板.docx
│   ├── 水利_模板.docx
│   ├── 市政房建_模板.docx
│   ├── 其他_模板.docx
│   └── 招标公告_模板.docx
│
├── output/                  # 📤 生成结果（自动创建）
│   ├── 公路/
│   ├── 水利/
│   ├── 市政房建/
│   └── 其他/
│
└── .github/
    └── workflows/
        └── generate-bid-docs.yml  # 🤖 GitHub Actions 工作流
```

---

## 🚀 快速开始

### 第一步：安装依赖

```bash
pip install -r requirements.txt
```

### 第二步：准备模板

从您现有的样本文件初始化模板：

```bash
python init_templates.py \
  --公路    "C:/Desktop/新建文件夹/公路/1.沙厦高速公路...标段.doc" \
  --水利    "C:/Desktop/新建文件夹/水利/1.河龙贡米...工程.doc" \
  --市政房建 "C:/Desktop/新建文件夹/市政房建/1.沙县区...项目.doc" \
  --其他    "C:/Desktop/新建文件夹/其他/1.沙县小吃...监理.doc" \
  --notice  "C:/Desktop/招标D/福建省房屋建筑和市政基础设施工程标准施工.doc"
```

> 📝 脚本会将文件复制到 `templates/` 目录，然后**在 Word 中手动将关键字段替换为占位符**（如 `{{PROJECT_NAME}}`）。

### 第三步：编辑项目配置

打开 `projects.json`，填写各项目信息：

```json
{
  "projects": [
    {
      "INDEX": "1",
      "TYPE": "公路",
      "PROJECT_NAME": "您的项目名称",
      "PROJECT_CODE": "GCZB-2025-GJ-001",
      "OWNER_NAME": "招标人单位名称",
      ...
    }
  ]
}
```

### 第四步：运行生成

```bash
python generate_bid.py --config projects.json --output ./output
```

---

## 📋 占位符说明

在模板文件中，使用 `{{字段名}}` 格式标记需要动态替换的内容：

| 占位符 | 说明 | 示例 |
|--------|------|------|
| `{{PROJECT_NAME}}` | 项目名称 | 沙厦高速公路沙县虬江互通... |
| `{{PROJECT_CODE}}` | 项目编号 | GCZB-2025-GJ-001 |
| `{{OWNER_NAME}}` | 招标人名称 | 福建省沙厦高速公路建设有限公司 |
| `{{OWNER_ADDRESS}}` | 招标人地址 | 三明市沙县区政府路1号 |
| `{{OWNER_CONTACT}}` | 招标人联系方式 | 0598-12345678 |
| `{{AGENT_NAME}}` | 招标代理机构 | 福建XX招标代理有限公司 |
| `{{AGENT_ADDRESS}}` | 代理机构地址 | 三明市梅列区XX路 |
| `{{AGENT_CONTACT}}` | 代理机构联系 | 0598-87654321 |
| `{{BID_SECTION}}` | 标段名称 | 机电、交安、绿化标段 |
| `{{BID_AMOUNT}}` | 估算金额 | 3500万元 |
| `{{QUALIFY_LEVEL}}` | 资质等级要求 | 公路工程专业承包二级及以上 |
| `{{DEADLINE_DATE}}` | 投标截止时间 | 2025年5月20日9时00分 |
| `{{OPEN_DATE}}` | 开标时间 | 2025年5月20日9时30分 |
| `{{OPEN_LOCATION}}` | 开标地点 | 三明市公共资源交易中心 |
| `{{SELL_START_DATE}}` | 文件发售开始 | 2025年4月28日 |
| `{{SELL_END_DATE}}` | 文件发售截止 | 2025年5月15日 |
| `{{BOND_AMOUNT}}` | 投标保证金 | 20万元 |
| `{{CONTRACT_PERIOD}}` | 合同工期（天） | 240 |
| `{{QUALITY_LEVEL}}` | 质量标准 | 合格 |
| `{{LOCATION}}` | 工程地点 | 福建省三明市沙县区 |
| `{{YEAR}}` | 当前年份（自动） | 2025 |
| `{{MONTH}}` | 当前月份（自动） | 04 |
| `{{DAY}}` | 当前日期（自动） | 28 |

---

## 🤖 GitHub Actions 自动化

### 工作流说明

工作流文件：`.github/workflows/generate-bid-docs.yml`

**触发方式：**

| 触发场景 | 说明 |
|----------|------|
| 手动触发 | 在 GitHub Actions 页面点击 "Run workflow" |
| 自动触发 | 推送 `projects.json` 或模板文件到 main 分支时 |
| 定时触发 | 每天 08:00（北京时间）自动检查并生成 |

**工作流阶段：**

```
validate（验证）→ generate（生成）→ release（发布）→ notify（通知）
```

1. **validate**：检查 JSON 格式、字段完整性、模板是否存在
2. **generate**：安装依赖、执行生成脚本、上传为 Artifact（保留30天）
3. **release**：打包 ZIP 并创建 GitHub Release，持久保存
4. **notify**：输出最终执行状态汇总

### 本地测试工作流

```bash
# 安装 act（本地运行 GitHub Actions）
# https://github.com/nektos/act

act workflow_dispatch -W .github/workflows/generate-bid-docs.yml
```

---

## 📂 工程类型说明

| 类型 | 使用模板 | 适用场景 |
|------|----------|----------|
| `公路` | 公路_模板.docx | 道路、高速公路、桥梁等工程 |
| `水利` | 水利_模板.docx | 水利、农田、生态治理、小流域治理等 |
| `市政房建` | 市政房建_模板.docx | 市政基础设施、污水处理、厂房等 |
| `其他` | 其他_模板.docx | 监理服务、全过程咨询、勘察设计等 |

---

## ❓ 常见问题

**Q: .doc 文件无法转换为 .docx？**  
A: 请安装 LibreOffice，或手动在 Word 中另存为 `.docx` 格式后放入 `templates/` 目录。

**Q: 占位符没有被替换？**  
A: 检查模板中的占位符格式是否为 `{{FIELD_NAME}}`（双花括号），确保字段名与 `projects.json` 中的 key 完全一致。

**Q: 如何新增工程类型？**  
A: 在 `templates/` 中添加 `新类型_模板.docx`，在 `projects.json` 的 `TYPE` 字段填写对应类型名称即可。

**Q: 如何批量处理大量项目？**  
A: 在 `projects.json` 的 `projects` 数组中添加更多对象，每个对象对应一个项目，脚本会自动批量生成。

---

## 📞 技术支持

如需扩展功能（如 PDF 输出、电子签章、在线填报表单），请在项目 Issue 中提出需求。
