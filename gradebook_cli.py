#!/usr/bin/env python3
"""CLI entry point for the gradebook tool."""

import sys
from pathlib import Path

import click

from core import load_config, process


@click.command()
@click.option(
    "--input", "-i", "input_dir",
    required=True,
    type=click.Path(exists=True, file_okay=False, dir_okay=True, path_type=Path),
    help="Folder containing MS Teams .xlsx export files.",
)
@click.option(
    "--config", "-c", "config_path",
    default=None,
    type=click.Path(exists=True, path_type=Path),
    help="Path to config.json / config.yaml (defaults to config.json in --input folder).",
)
@click.option(
    "--output", "-o", "output_path",
    default="gradebook.xlsx",
    type=click.Path(path_type=Path),
    show_default=True,
    help="Output .xlsx file path.",
)
@click.option(
    "--fuzzy-threshold", default=60, show_default=True,
    help="Minimum fuzzy-match score (0-100) to assign an export to a category.",
)
def main(input_dir: Path, config_path: Path | None, output_path: Path, fuzzy_threshold: int):
    """
    Build a weighted gradebook .xlsx from MS Teams assignment exports.

    \b
    Example:
        python gradebook.py --input ./exports --config config.json --output gradebook.xlsx
    """
    # Locate config
    if config_path is None:
        for candidate in ["config.json", "config.yaml", "config.yml"]:
            p = input_dir / candidate
            if p.exists():
                config_path = p
                break
        if config_path is None:
            click.echo(
                "❌  No config file found. Pass --config or place config.json in the input folder.",
                err=True,
            )
            sys.exit(1)

    click.echo(f"📂  Input folder : {input_dir}")
    click.echo(f"⚙️   Config       : {config_path}")

    config = load_config(config_path)
    click.echo(f"📘  Course        : {config.get('course', '(unnamed)')}")
    if config.get("professors"):
        click.echo(f"👩‍🏫  Instructors  : {', '.join(config['professors'])}")

    # Collect xlsx files
    sources = sorted(input_dir.glob("*.xlsx"))
    if not sources:
        click.echo(f"❌  No .xlsx files found in {input_dir}", err=True)
        sys.exit(1)

    click.echo(f"\n📄  Found {len(sources)} export file(s):")
    for s in sources:
        click.echo(f"    • {s.name}")

    click.echo("\n⏳  Processing …")
    try:
        xlsx_bytes, unmatched = process(sources, config)
    except Exception as exc:
        click.echo(f"\n❌  Error: {exc}", err=True)
        sys.exit(1)

    if unmatched:
        click.echo("\n⚠️   Unmatched assignments (added to 'Unassigned' column):")
        for name in unmatched:
            click.echo(f"    • {name}")

    output_path.write_bytes(xlsx_bytes)
    click.echo(f"\n✅  Gradebook written to: {output_path}")


if __name__ == "__main__":
    main()
