"""
main.py — CLI 入口（完整版）
"""
import os
import sys
import glob
import time
import click
from rich.console import Console
from rich.table import Table

import config
from models import ParsedDocument
from parsers.pdf_parser import PDFParser
from parsers.word_parser import WordParser
from parsers.engine import ProtocolEngine
from excel_writer import ExcelWriter

console = Console()


def get_parser(filepath: str):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".pdf":
        return PDFParser(filepath)
    elif ext in (".docx", ".doc"):
        return WordParser(filepath)
    else:
        return None


def process_single(filepath: str, output_path: str = None) -> str:
    filename = os.path.basename(filepath)
    name_no_ext = os.path.splitext(filename)[0]

    parser = get_parser(filepath)
    if parser is None:
        console.print(f"[red]  ✗ 不支持的文件格式: {filename}[/red]")
        return ""

    console.print(f"  [cyan]📄 解析文档...[/cyan]")
    doc = parser.parse()

    console.print(f"  [cyan]🔍 智能提取点位...[/cyan]")
    engine = ProtocolEngine()
    doc = engine.process(doc)

    if len(doc.points) == 0:
        console.print(f"  [yellow]⚠ 未提取到点位数据，请检查文档格式[/yellow]")

    if output_path is None:
        output_path = os.path.join(config.OUTPUT_DIR, f"{name_no_ext}_标准点位表.xlsx")

    console.print(f"  [cyan]📊 生成Excel...[/cyan]")
    writer = ExcelWriter(doc)
    result = writer.write(output_path)
    return result


def show_result_table(results: list):
    table = Table(title="🎉 处理完成", show_lines=True)
    table.add_column("序号", justify="center", style="bold", width=5)
    table.add_column("输入文件", style="cyan", max_width=35)
    table.add_column("输出文件", style="green", max_width=45)
    table.add_column("状态", justify="center", width=8)

    for i, r in enumerate(results, 1):
        status = "[green]✓ 成功[/green]" if r["ok"] else "[red]✗ 失败[/red]"
        table.add_row(
            str(i),
            r["input"],
            os.path.basename(r["output"]) if r["output"] else "-",
            status,
        )
    console.print()
    console.print(table)


@click.command()
@click.option(
    "-i", "--input", "input_files",
    multiple=True,
    help="输入文件路径（可多个），不指定则处理 input/ 下所有 PDF/Word",
)
@click.option(
    "-o", "--output", "output_path",
    default=None,
    help="输出Excel路径（仅单文件时有效）",
)
def main(input_files, output_path):
    """协议文档 → 标准Excel 自动化转换工具"""
    console.print()
    console.print("[bold blue]╔══════════════════════════════════════╗[/bold blue]")
    console.print("[bold blue]║   协议文档 → 标准Excel 转换工具     ║[/bold blue]")
    console.print("[bold blue]╚══════════════════════════════════════╝[/bold blue]")
    console.print()

    # 收集输入文件
    files = []
    if input_files:
        for f in input_files:
            if os.path.isfile(f):
                files.append(os.path.abspath(f))
            else:
                console.print(f"[red]文件不存在: {f}[/red]")
    else:
        os.makedirs(config.INPUT_DIR, exist_ok=True)
        for ext in ("*.pdf", "*.PDF", "*.docx", "*.DOCX", "*.doc"):
            files.extend(glob.glob(os.path.join(config.INPUT_DIR, ext)))

    if not files:
        console.print("[yellow]未找到待处理文件。[/yellow]")
        console.print(f"[dim]请将 PDF/Word 文件放入 {config.INPUT_DIR}/ 目录[/dim]")
        return

    files = sorted(set(files))
    console.print(f"[green]找到 {len(files)} 个文件待处理[/green]")
    console.print()

    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    results = []

    for idx, fpath in enumerate(files, 1):
        fname = os.path.basename(fpath)
        console.print(f"[bold]━━━ [{idx}/{len(files)}] {fname} ━━━[/bold]")
        t0 = time.time()

        try:
            out = output_path if (output_path and len(files) == 1) else None
            result_path = process_single(fpath, out)
            elapsed = time.time() - t0

            if result_path:
                console.print(
                    f"  [green]✓ 完成 ({elapsed:.1f}s) → {os.path.basename(result_path)}[/green]"
                )
                results.append({"input": fname, "output": result_path, "ok": True})
            else:
                console.print(f"  [red]✗ 处理失败[/red]")
                results.append({"input": fname, "output": "", "ok": False})

        except Exception as e:
            console.print(f"  [red]✗ 异常: {e}[/red]")
            results.append({"input": fname, "output": "", "ok": False})

        console.print()

    show_result_table(results)

    success = sum(1 for r in results if r["ok"])
    console.print()
    console.print(
        f"[bold green]成功: {success}/{len(results)}[/bold green]  "
        f"输出目录: [cyan]{os.path.abspath(config.OUTPUT_DIR)}[/cyan]"
    )
    console.print()


if __name__ == "__main__":
    main()
