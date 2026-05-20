from __future__ import annotations

import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any


TABLE_PATTERN = re.compile(r"(?i)\b(?:from|join)\s+([`\w\.]+)")
PARAM_PATTERN = re.compile(r"\^P_PARAM\.([^^]+)\^")
JOIN_PATTERN = re.compile(
    r"(?i)\b([a-zA-Z_][\w]*)\.([a-zA-Z_][\w]*)\s*=\s*([a-zA-Z_][\w]*)\.([a-zA-Z_][\w]*)"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="根据 SmartBI meta 与 SQL 自动生成数据集知识库初稿。")
    parser.add_argument(
        "--requirement-dir",
        default="",
        help="需求文件夹路径。传入后，若未单独指定输出路径，则默认输出到该目录下。",
    )
    parser.add_argument(
        "--project-dir",
        default="",
        help="ai_dict_project 根目录。为空时自动按脚本所在位置向上推断。",
    )
    parser.add_argument(
        "--meta-path",
        default="data/smartbi/meta/smartbi_report_meta.json",
        help="报表元数据 JSON 路径，默认 data/smartbi/meta/smartbi_report_meta.json。",
    )
    parser.add_argument(
        "--business-kb-path",
        default="",
        help="业务口径知识库路径，可选。",
    )
    parser.add_argument(
        "--output-path",
        default="",
        help="输出知识库路径，可选。为空时默认写入 knowledge_base/knowledge_base_<scope>_dataset.json。",
    )
    parser.add_argument(
        "--domain-name",
        default="",
        help="板块中文名，可选。不传时优先取 target_root 或 report_level_1。",
    )
    parser.add_argument(
        "--scope",
        default="",
        help="知识库 scope，例如 overseas_subject。可选。",
    )
    parser.add_argument(
        "--top-table-limit",
        type=int,
        default=20,
        help="高频表保留数量，默认 20。",
    )
    parser.add_argument(
        "--top-join-limit",
        type=int,
        default=20,
        help="高频 join 提示保留数量，默认 20。",
    )
    return parser.parse_args()


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def normalize_scope(value: str) -> str:
    cleaned = re.sub(r"[^\w]+", "_", value.strip().lower(), flags=re.UNICODE).strip("_")
    return cleaned or "dataset_domain"


def resolve_project_dir(cli_value: str) -> Path:
    if cli_value.strip():
        return Path(cli_value).resolve()
    script_dir = Path(__file__).resolve().parent
    if script_dir.name == "需求实施文档":
        return script_dir.parent
    return script_dir


def collect_sql_stats(sql_path: Path) -> dict[str, Any]:
    text = sql_path.read_text(encoding="utf-8", errors="ignore")

    tables: list[str] = []
    for match in TABLE_PATTERN.finditer(text):
        name = match.group(1).strip("`")
        if "." not in name:
            continue
        tables.append(name)

    params = sorted({match.group(1) for match in PARAM_PATTERN.finditer(text)})

    joins: list[dict[str, str]] = []
    for match in JOIN_PATTERN.finditer(text):
        joins.append(
            {
                "left_alias": match.group(1),
                "left_column": match.group(2),
                "right_alias": match.group(3),
                "right_column": match.group(4),
            }
        )

    return {
        "tables": tables,
        "params": params,
        "joins": joins,
    }


def build_folder_distribution(meta_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    counter: Counter[str] = Counter()
    for item in meta_items:
        key = " / ".join(item.get("folder_path") or ["根目录"])
        counter[key] += 1
    return [{"folder": name, "report_count": count} for name, count in counter.most_common()]


def infer_domain_name(meta_items: list[dict[str, Any]], cli_value: str) -> str:
    if cli_value.strip():
        return cli_value.strip()
    for key in ("target_root", "report_level_1"):
        values = [str(item.get(key, "")).strip() for item in meta_items if str(item.get(key, "")).strip()]
        if values:
            return Counter(values).most_common(1)[0][0]
    return "未命名板块"


def resolve_output_path(project_dir: Path, requirement_dir: Path | None, scope: str, cli_output: str) -> Path:
    if cli_output.strip():
        return (project_dir / cli_output).resolve() if not Path(cli_output).is_absolute() else Path(cli_output)
    if requirement_dir is not None:
        return requirement_dir / f"{scope}_dataset_kb.json"
    return project_dir / "knowledge_base" / f"knowledge_base_{scope}_dataset.json"


def resolve_sql_path(project_dir: Path, raw_sql_file: str) -> Path:
    sql_path = Path(raw_sql_file)
    if sql_path.exists():
        return sql_path

    candidate = project_dir / raw_sql_file
    if candidate.exists():
        return candidate

    path_text = raw_sql_file.replace("/", "\\")
    anchor = "\\海外直播业务线\\"
    if anchor in path_text:
        suffix = path_text.split(anchor, 1)[1]
        candidate = project_dir / "海外直播业务线" / Path(suffix)
        if candidate.exists():
            return candidate

    return sql_path


def build_dataset_inventory(project_dir: Path, meta_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    inventory: list[dict[str, Any]] = []
    for item in meta_items:
        folder_path = " / ".join(item.get("folder_path") or ["根目录"])
        for dataset in item.get("datasets", []):
            sql_file = str(dataset.get("sql_file", "")).replace(str(project_dir) + "\\", "")
            inventory.append(
                {
                    "report_name": item.get("report_name", ""),
                    "report_id": item.get("report_id", ""),
                    "report_group": folder_path,
                    "report_path": item.get("report_path", ""),
                    "dataset_name": dataset.get("dataset_name", ""),
                    "dataset_id": dataset.get("dataset_id", ""),
                    "dataset_type": dataset.get("dataset_type", ""),
                    "sql_file": sql_file,
                }
            )
    return inventory


def main() -> None:
    args = parse_args()
    project_dir = resolve_project_dir(args.project_dir)
    requirement_dir = Path(args.requirement_dir).resolve() if args.requirement_dir else None
    meta_path = project_dir / args.meta_path
    if not meta_path.exists():
        raise FileNotFoundError(f"未找到报表元数据文件: {meta_path}")

    meta_items = read_json(meta_path)
    if not isinstance(meta_items, list):
        raise ValueError(f"报表元数据格式不正确，期望 list，实际为: {type(meta_items).__name__}")

    domain_name = infer_domain_name(meta_items, args.domain_name)
    scope = normalize_scope(args.scope or domain_name)
    output_path = resolve_output_path(project_dir, requirement_dir, scope, args.output_path)

    table_counter: Counter[str] = Counter()
    table_reports: defaultdict[str, set[str]] = defaultdict(set)
    table_sql_files: defaultdict[str, set[str]] = defaultdict(set)
    param_counter: Counter[str] = Counter()
    param_reports: defaultdict[str, set[str]] = defaultdict(set)
    param_sql_files: defaultdict[str, set[str]] = defaultdict(set)
    join_counter: Counter[tuple[str, str]] = Counter()
    join_reports: defaultdict[tuple[str, str], set[str]] = defaultdict(set)
    scanned_sql_files: list[str] = []
    missing_sql_files: list[str] = []

    for item in meta_items:
        report_name = str(item.get("report_name", "")).strip()
        for dataset in item.get("datasets", []):
            raw_sql_file = str(dataset.get("sql_file", "")).strip()
            if not raw_sql_file:
                continue

            sql_path = resolve_sql_path(project_dir, raw_sql_file)

            if not sql_path.exists():
                missing_sql_files.append(raw_sql_file)
                continue

            scanned_sql_files.append(str(sql_path))
            stats = collect_sql_stats(sql_path)

            for table_name in stats["tables"]:
                table_counter[table_name] += 1
                table_reports[table_name].add(report_name)
                table_sql_files[table_name].add(str(sql_path))

            for param_name in stats["params"]:
                param_counter[param_name] += 1
                param_reports[param_name].add(report_name)
                param_sql_files[param_name].add(str(sql_path))

            for join_item in stats["joins"]:
                join_key = tuple(sorted((join_item["left_column"], join_item["right_column"])))
                join_counter[join_key] += 1
                join_reports[join_key].add(report_name)

    common_tables: list[dict[str, Any]] = []
    for table_name, usage_count in table_counter.most_common(args.top_table_limit):
        common_tables.append(
            {
                "table_name": table_name,
                "usage_count": usage_count,
                "role": "待补充：该表在本板块中的业务角色。",
                "common_join_keys": [],
                "common_filters": [],
                "referenced_by_reports": sorted(table_reports[table_name])[:20],
                "sample_sql_files": sorted(table_sql_files[table_name])[:5],
            }
        )

    relationship_hints: list[dict[str, Any]] = []
    for join_key, usage_count in join_counter.most_common(args.top_join_limit):
        relationship_hints.append(
            {
                "join_columns": list(join_key),
                "usage_count": usage_count,
                "description": "待补充：说明这组字段通常承载什么关联关系。",
                "referenced_by_reports": sorted(join_reports[join_key])[:20],
            }
        )

    parameter_catalogue: list[dict[str, Any]] = []
    for param_name, usage_count in param_counter.most_common():
        parameter_catalogue.append(
            {
                "smartbi_param": f"^P_PARAM.{param_name}^",
                "param_name": param_name.split(".")[-1],
                "usage_count": usage_count,
                "description": "待补充：该参数的业务含义、是过滤还是展示层级控制。",
                "referenced_by_reports": sorted(param_reports[param_name])[:20],
                "sample_sql_files": sorted(param_sql_files[param_name])[:5],
            }
        )

    business_kb_reference: dict[str, Any] = {}
    if args.business_kb_path.strip():
        business_kb_path = (
            project_dir / args.business_kb_path
            if not Path(args.business_kb_path).is_absolute()
            else Path(args.business_kb_path)
        )
        business_kb_reference = {
            "file": str(business_kb_path),
            "usage": "业务口径知识库由人工维护，建议与数据集知识库联合使用。",
        }

    dataset_inventory = build_dataset_inventory(project_dir, meta_items)
    report_groups = build_folder_distribution(meta_items)

    payload: dict[str, Any] = {
        "project": "ai_dict_project",
        "scope": scope,
        "domain_name": domain_name,
        "knowledge_base_type": "dataset_knowledge_base",
        "description": "由脚本自动生成的数据集知识库初稿。适合作为板块级数据集地图、表使用清单、参数目录和后续人工补充的基础。",
        "source_summary": {
            "source_meta_file": str(meta_path),
            "report_count": len(meta_items),
            "dataset_count": sum(int(item.get("dataset_count", 0) or 0) for item in meta_items),
            "folder_distribution": report_groups,
            "scanned_sql_file_count": len(scanned_sql_files),
            "missing_sql_files": missing_sql_files,
        },
        "business_kb_reference": business_kb_reference,
        "common_tables": common_tables,
        "table_relationship_hints": relationship_hints,
        "parameter_catalogue": parameter_catalogue,
        "common_filter_patterns": [
            "待补充：沉淀本板块的通用过滤条件，例如有效课节、已签到、品牌、订单状态等。",
            "待补充：区分真正过滤数据的参数，和只控制展示层级的参数。",
        ],
        "field_source_rules": [
            {
                "field_or_metric": "示例指标",
                "source_rule": "待补充：字段或指标从哪张表、哪个字段来，是否需要中间加工。",
            }
        ],
        "dataset_patterns": [
            {
                "pattern_name": "待补充：本板块高频数据集模式",
                "applies_to": [],
                "characteristics": [],
            }
        ],
        "implementation_constraints": [
            "待补充：本板块的特殊限制条件、边界口径、性能风险、关键词规则或目标表依赖。",
        ],
        "field_lineage_templates": [],
        "parameter_mapping_templates": [],
        "dataset_inventory": dataset_inventory,
        "notes": [
            "该文件由 build_dataset_knowledge_base.py 自动生成。",
            "自动脚本擅长整理表清单、参数目录、SQL 使用痕迹；业务角色、指标含义、精确血缘仍建议人工二次补充。",
            "建议后续在 common_tables / parameter_catalogue / field_lineage_templates / parameter_mapping_templates 四块持续迭代。",
        ],
    }

    write_json(output_path, payload)
    print(f"已生成数据集知识库: {output_path}")


if __name__ == "__main__":
    main()
