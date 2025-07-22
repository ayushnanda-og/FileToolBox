# app.py
import typer
from pathlib import Path
from pdf_tools import merge_pdfs, split_pdf

app = typer.Typer(help="FileToolbox: Merge & Split PDF Files")

@app.command()
def merge(
    inputs: list[Path] = typer.Argument(..., help="PDF files to merge"),
    output: Path = typer.Option(..., "--output", "-o", help="Output PDF path")
):
    """Merge multiple PDFs into one."""
    result = merge_pdfs(inputs, output)
    typer.echo(f"Merged into: {result}")

@app.command()
def split(
    source: Path = typer.Argument(..., help="PDF file to split"),
    outdir: Path = typer.Option("output", "--outdir", "-d", help="Directory for split PDFs")
):
    """Split each page into a separate PDF."""
    results = split_pdf(source, outdir)
    typer.echo(f"Split into {len(results)} files in {outdir}")

if __name__ == "__main__":
    app()
