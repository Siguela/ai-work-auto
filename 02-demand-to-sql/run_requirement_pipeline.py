from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按板块级知识库 + 当前需求输入，生成当前需求的实施文档。"
    )
    parser.add_argument("--requirement-dir", required=True, help="需求文件夹路径。")
    parser.add_argument("--project-dir", default="", help="ai_dict_project 根目录。为空时自动推断。")
    parser.add_argument("--scope", required=True, help="板块 scope，例如 overseas_subject。")
    parser.add_argument("--domain-name", default="", help="板块中文名，例如 海外学科。")
    parser.add_argument(
        "--meta-path",
        default="data/smartbi/meta/smartbi_report_meta.json",
        help="仅在刷新数据集知识库时使用的报表元数据路径，默认 data/smartbi/meta/smartbi_report_meta.json。",
    )
    parser.add_argument("--business-kb-path", default="", help="业务口径知识库路径。为空时默认读取 knowledge_base/knowledge_base_<scope>.json。")
    parser.add_argument("--request-json", default="", help="需求输入文件路径，可选。可传 JSON 或 Markdown。")
    parser.add_argument("--dataset-kb-path", default="", help="数据集知识库路径。为空时默认读取 knowledge_base/knowledge_base_<scope>_dataset.json。")
    parser.add_argument("--doc-output-path", default="", help="需求实施文档输出路径，可选。")
    parser.add_argument("--refresh-dataset-kb", action="store_true", help="显式刷新板块数据集知识库。默认不刷新，直接读取现有知识库。")
    return parser.parse_args()


def resolve_project_dir(cli_value: str) -> Path:
    if cli_value.strip():
        return Path(cli_value).resolve()
    script_dir = Path(__file__).resolve().parent
    if script_dir.name == "需求实施文档":
        return script_dir.parent
    return script_dir


def resolve_path(base_dir: Path, value: str) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path.resolve()
    return (base_dir / path).resolve()


def resolve_board_kb_path(project_dir: Path, scope: str, cli_value: str, suffix: str) -> Path:
    if cli_value.strip():
        path = Path(cli_value)
        if path.is_absolute():
            return path.resolve()
        return (project_dir / path).resolve()
    return (project_dir / "knowledge_base" / f"knowledge_base_{scope}{suffix}").resolve()


def run_command(cmd: list[str], title: str) -> None:
    print(f"\n[{title}]")
    print(" ".join(f'"{part}"' if " " in part else part for part in cmd))
    subprocess.run(cmd, check=True)


def main() -> None:
    args = parse_args()
    requirement_dir = Path(args.requirement_dir).resolve()
    if not requirement_dir.exists():
        raise FileNotFoundError(f"未找到需求文件夹: {requirement_dir}")

    project_dir = resolve_project_dir(args.project_dir)
    script_dir = Path(__file__).resolve().parent
    build_script = script_dir / "build_dataset_knowledge_base.py"
    doc_script = script_dir / "generate_requirement_implementation_doc.py"

    if not doc_script.exists():
        raise FileNotFoundError(f"未找到实施文档生成脚本: {doc_script}")

    dataset_kb_path = resolve_board_kb_path(
        project_dir,
        args.scope,
        args.dataset_kb_path,
        "_dataset.json",
    )
    business_kb_path = resolve_board_kb_path(
        project_dir,
        args.scope,
        args.business_kb_path,
        ".json",
    )

    if args.refresh_dataset_kb:
        if not build_script.exists():
            raise FileNotFoundError(f"未找到知识库生成脚本: {build_script}")
        kb_cmd = [
            sys.executable,
            str(build_script),
            "--project-dir",
            str(project_dir),
            "--scope",
            args.scope,
            "--meta-path",
            args.meta_path,
            "--output-path",
            str(dataset_kb_path),
        ]
        if args.domain_name.strip():
            kb_cmd.extend(["--domain-name", args.domain_name])
        if business_kb_path.exists():
            kb_cmd.extend(["--business-kb-path", str(business_kb_path)])
        run_command(kb_cmd, "步骤1/2 刷新板块数据集知识库")

    if not dataset_kb_path.exists():
        raise FileNotFoundError(
            f"未找到板块数据集知识库: {dataset_kb_path}\n"
            f"请先准备该文件，或加 --refresh-dataset-kb 先刷新一版。"
        )
    if not business_kb_path.exists():
        raise FileNotFoundError(
            f"未找到板块业务口径知识库: {business_kb_path}\n"
            f"请检查 --business-kb-path 或 knowledge_base 目录下的文件命名。"
        )

    request_input_path = resolve_path(requirement_dir, args.request_json) if args.request_json.strip() else None
    doc_output_path = resolve_path(requirement_dir, args.doc_output_path) if args.doc_output_path.strip() else None

    doc_cmd = [
        sys.executable,
        str(doc_script),
        "--requirement-dir",
        str(requirement_dir),
        "--dataset-kb-path",
        str(dataset_kb_path),
        "--business-kb-path",
        str(business_kb_path),
    ]
    if request_input_path is not None:
        doc_cmd.extend(["--request-json", str(request_input_path)])
    if doc_output_path is not None:
        doc_cmd.extend(["--output-path", str(doc_output_path)])

    run_command(doc_cmd, "步骤1/1 生成需求实施文档")
    print("\n已完成：读取板块知识库 + 当前需求输入，并生成需求实施文档。")


if __name__ == "__main__":
    main()
