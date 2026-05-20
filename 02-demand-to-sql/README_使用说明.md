# README 使用说明

## 1. 这套工具是做什么的

这套工具用于两件事：

1. 根据 SmartBI 报表元数据和现有 SQL，自动生成某个板块的“数据集知识库”初稿
2. 根据“需求输入 JSON + 数据集知识库 + 业务口径知识库”，自动生成“需求实施文档”草稿

适用场景：

- 海外学科
- 海外教务
- 海外产运
- 海外后端
- 其他沿用 `ai_dict_project` 目录结构的板块

## 2. 文件说明

当前主要文件如下：

- [build_dataset_knowledge_base.py](C:\wxy\海外学科数据字典\source_code\ai_dict_project\build_dataset_knowledge_base.py)
  作用：生成数据集知识库初稿
- [generate_requirement_implementation_doc.py](C:\wxy\海外学科数据字典\source_code\ai_dict_project\generate_requirement_implementation_doc.py)
  作用：生成需求实施文档草稿
- [run_requirement_pipeline.py](C:\wxy\海外学科数据字典\source_code\ai_dict_project\需求实施文档\run_requirement_pipeline.py)
  作用：一键串行执行“生成知识库 + 生成需求实施文档”
- [templates/requirement_request_template.json](C:\wxy\海外学科数据字典\source_code\ai_dict_project\templates\requirement_request_template.json)
  作用：需求输入模板
- [knowledge_base/knowledge_base_overseas_subject.json](C:\wxy\海外学科数据字典\source_code\ai_dict_project\knowledge_base\knowledge_base_overseas_subject.json)
  作用：业务口径知识库示例
- [knowledge_base/knowledge_base_overseas_subject_dataset.json](C:\wxy\海外学科数据字典\source_code\ai_dict_project\knowledge_base\knowledge_base_overseas_subject_dataset.json)
  作用：数据集知识库示例
- [prompts/requirement_implementation_template.md](C:\wxy\海外学科数据字典\source_code\ai_dict_project\prompts\requirement_implementation_template.md)
  作用：需求实施文档结构模板

## 3. 使用前要准备什么

需要具备以下内容：

1. 一个完整的 `ai_dict_project` 项目目录
2. `data/smartbi/meta/smartbi_report_meta.json`
3. 报表对应的 `.sql` 文件已经落盘
4. Python 3.10 及以上

建议目录结构如下：

```text
ai_dict_project/
├─ data/
│  └─ smartbi/meta/smartbi_report_meta.json
├─ knowledge_base/
├─ outputs/
├─ prompts/
├─ templates/
├─ 海外直播业务线/
│  └─ 某板块/
│     └─ 各报表目录/*.sql
├─ build_dataset_knowledge_base.py
└─ generate_requirement_implementation_doc.py
```

## 4. 第一步：生成数据集知识库

### 4.1 最简单的命令

在项目根目录执行：

```powershell
python build_dataset_knowledge_base.py `
  --project-dir "C:\你的项目\ai_dict_project" `
  --scope "overseas_subject" `
  --domain-name "海外学科"
```

如果你已经按“一个需求一个文件夹”的方式管理，推荐直接这样执行：

```powershell
python 需求实施文档\build_dataset_knowledge_base.py `
  --requirement-dir "C:\你的项目\ai_dict_project\需求实施文档\某个具体需求" `
  --scope "overseas_subject" `
  --domain-name "海外学科"
```

### 4.2 常用完整命令

```powershell
python build_dataset_knowledge_base.py `
  --project-dir "C:\你的项目\ai_dict_project" `
  --scope "overseas_subject" `
  --domain-name "海外学科" `
  --meta-path "data\smartbi\meta\smartbi_report_meta.json" `
  --business-kb-path "knowledge_base\knowledge_base_overseas_subject.json" `
  --output-path "knowledge_base\knowledge_base_overseas_subject_dataset.json"
```

### 4.3 参数说明

- `--project-dir`
  说明：`ai_dict_project` 根目录
- `--scope`
  说明：知识库英文标识，建议使用类似 `overseas_subject`
- `--domain-name`
  说明：板块中文名，例如 `海外学科`
- `--meta-path`
  说明：报表元数据路径，默认 `data/smartbi/meta/smartbi_report_meta.json`
- `--business-kb-path`
  说明：业务口径知识库路径，可选
- `--output-path`
  说明：输出路径，可选；不传时默认写入 `knowledge_base/knowledge_base_<scope>_dataset.json`

### 4.4 生成后会得到什么

会生成一份 JSON 文件，主要包含：

- 报表和数据集清单
- 高频底表统计
- 参数目录
- join 关系提示
- 可继续人工补充的数据集知识库骨架

注意：

- 这一步生成的是“初稿”
- 脚本擅长自动整理结构和痕迹
- 业务角色、精确字段血缘、口径说明仍建议人工补充

## 5. 第二步：准备需求输入 JSON

复制模板文件：

- [templates/requirement_request_template.json](C:\wxy\海外学科数据字典\source_code\ai_dict_project\templates\requirement_request_template.json)

按实际需求修改以下字段：

- `requirement_name`
- `domain_name`
- `business_background`
- `goal`
- `candidate_report_names`
- `core_dimensions`
- `core_metrics`
- `filters`
- `special_calibers`
- `mandatory_filters`
- `validation_plan`
- `open_questions`

## 6. 第三步：生成需求实施文档

### 6.1 最简单的命令

```powershell
python generate_requirement_implementation_doc.py `
  --request-json "C:\你的项目\ai_dict_project\templates\requirement_request_template.json" `
  --dataset-kb-path "C:\你的项目\ai_dict_project\knowledge_base\knowledge_base_overseas_subject_dataset.json" `
  --output-path "C:\你的项目\ai_dict_project\outputs\requirement_doc_sample.md"
```

如果需求文件夹里已经放好了：

- `requirement_request.json` 或 `需求输入.json`
- `*dataset_kb.json`

那么可以直接只传需求文件夹：

```powershell
python 需求实施文档\generate_requirement_implementation_doc.py `
  --requirement-dir "C:\你的项目\ai_dict_project\需求实施文档\某个具体需求"
```

默认行为：

- 自动读取需求文件夹内的 `requirement_request.json`、`需求输入.json` 或 `requirement_request_template.json`
- 自动读取需求文件夹内的 `*dataset_kb.json`
- 自动输出到需求文件夹下的 `需求实施文档.md`

### 6.2 常用完整命令

```powershell
python generate_requirement_implementation_doc.py `
  --request-json "C:\你的项目\ai_dict_project\templates\your_requirement.json" `
  --dataset-kb-path "C:\你的项目\ai_dict_project\knowledge_base\knowledge_base_overseas_subject_dataset.json" `
  --business-kb-path "C:\你的项目\ai_dict_project\knowledge_base\knowledge_base_overseas_subject.json" `
  --output-path "C:\你的项目\ai_dict_project\outputs\需求实施文档_示例.md"
```

### 6.3 生成后会得到什么

会生成一份 Markdown 文档，通常包含：

- 需求背景
- 需求范围
- 输出方案
- 主要事实表/维表
- 指标实现方案
- 参数映射关系
- 校验建议
- 给公司 AI 生成 SQL 的输入摘要

## 7. 推荐使用流程

建议按下面顺序使用：

1. 在 `需求实施文档` 目录下新建一个需求文件夹
2. 先准备好板块的 `smartbi_report_meta.json` 和 `.sql`
3. 运行 `build_dataset_knowledge_base.py`
4. 人工补充生成出来的数据集知识库
5. 将需求输入 JSON 放入该需求文件夹
6. 运行 `generate_requirement_implementation_doc.py`
7. 人工校对文档
8. 将第 13 节“给公司 AI 生成 SQL 的输入摘要”喂给公司 AI

## 7.1 一键串行执行

如果你希望一条命令完成“生成数据集知识库 + 生成需求实施文档”，可以直接执行：

```powershell
python 需求实施文档\run_requirement_pipeline.py `
  --requirement-dir "C:\你的项目\ai_dict_project\需求实施文档\某个具体需求" `
  --scope "overseas_subject" `
  --domain-name "海外学科" `
  --business-kb-path "knowledge_base\knowledge_base_overseas_subject.json"
```

说明：

- 第一步会先在需求文件夹内生成 `overseas_subject_dataset_kb.json`
- 第二步会自动读取需求文件夹内的需求输入 JSON，并生成 `需求实施文档.md`

如果只想跑其中一步：

- 只生成知识库：加 `--skip-doc`
- 只生成实施文档：加 `--skip-kb`

## 8. 推荐人工补充的位置

自动脚本生成后，建议优先补这几块：

- `common_tables[*].role`
- `common_tables[*].common_join_keys`
- `common_tables[*].common_filters`
- `field_source_rules`
- `dataset_patterns`
- `implementation_constraints`
- `field_lineage_templates`
- `parameter_mapping_templates`

这些内容越完整，后续生成需求实施文档就越准确。

## 9. 迁移给其他板块时怎么改

只需要改 4 类东西：

1. 板块目录
2. `smartbi_report_meta.json`
3. 业务口径知识库文件
4. 命令中的 `scope` 和 `domain-name`

例如迁移到“海外教务”：

```powershell
python build_dataset_knowledge_base.py `
  --project-dir "C:\你的项目\ai_dict_project" `
  --scope "overseas_academic_affairs" `
  --domain-name "海外教务" `
  --business-kb-path "knowledge_base\knowledge_base_overseas_academic_affairs.json" `
  --output-path "knowledge_base\knowledge_base_overseas_academic_affairs_dataset.json"
```

## 10. 常见问题

### 10.1 扫不到 SQL 文件

现象：

- 生成出来的知识库里 `scanned_sql_file_count = 0`
- `missing_sql_files` 很多

常见原因：

- `smartbi_report_meta.json` 里的 `sql_file` 还是旧机器的绝对路径
- 当前项目目录不完整
- `.sql` 文件没有复制过来

处理方式：

- 先确认 `海外直播业务线\板块\报表目录\*.sql` 是否真实存在
- 当前脚本已经兼容“旧绝对路径 + 新项目目录”的常见情况
- 如果目录结构变化太大，需要调整脚本中的路径解析逻辑

### 10.2 生成的知识库字段很多都是“待补充”

这是正常的。

原因：

- 脚本擅长自动提取结构
- 但“业务角色、口径解释、特殊规则”很难完全靠程序自动判断

正确做法：

- 把脚本结果当成“初稿”
- 再由熟悉业务的人做二次补充

### 10.3 需求实施文档里指标不够细

原因：

- 当前生成逻辑优先复用 `field_lineage_templates`
- 如果知识库里这部分还不完整，文档就会偏骨架化

建议：

- 先补齐对应板块的数据集知识库
- 尤其补 `field_lineage_templates` 和 `parameter_mapping_templates`

## 11. 适合交给同事的最小文件包

如果要迁移给同事，建议至少交付这些文件：

1. `build_dataset_knowledge_base.py`
2. `generate_requirement_implementation_doc.py`
3. `templates/requirement_request_template.json`
4. 一份该板块现有的业务口径知识库
5. `smartbi_report_meta.json`
6. 对应板块的 `.sql` 文件目录

## 12. 已验证样例

当前项目里已经验证过以下样例输出：

- [outputs/subject_auto_test_dataset_kb.json](C:\wxy\海外学科数据字典\source_code\ai_dict_project\outputs\subject_auto_test_dataset_kb.json)
- [outputs/requirement_doc_sample.md](C:\wxy\海外学科数据字典\source_code\ai_dict_project\outputs\requirement_doc_sample.md)

## 13. 建议的下一步

如果后续要继续增强，建议优先做这几件事：

1. 给 `build_dataset_knowledge_base.py` 增加“字段级血缘自动抽取”能力
2. 给 `generate_requirement_implementation_doc.py` 增加“按具体报表优先匹配”能力
3. 增加一个总入口脚本，一次性串起来跑“知识库生成 + 文档生成”
