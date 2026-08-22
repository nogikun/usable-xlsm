"""Check exported VBA sources for syntax errors before applying them."""

from pathlib import Path
import sys

import typer

from usable_xlsm import check_directory, check_file


# Excel's dialogs and VBA errors are localised, so force UTF-8 on the streams:
# otherwise Windows encodes them in the console codepage and any tool reading
# this output gets mojibake instead of the error message.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


app = typer.Typer(add_completion=False)


@app.command()
def check(
    input_path: Path = typer.Option(
        ...,
        "--input",
        "-i",
        exists=True,
        help="VBA source file or a directory of them.",
    ),
) -> None:
    """Report syntax errors, exiting non-zero when any are found."""
    issues = check_file(input_path) if input_path.is_file() else check_directory(input_path)

    if not issues:
        typer.echo("OK: no syntax errors found.")
        return

    for issue in issues:
        typer.echo(str(issue), err=True)
    typer.echo(f"\n{len(issues)} syntax error(s) found.", err=True)
    raise typer.Exit(code=1)


if __name__ == "__main__":
    app(args=["--help"] if len(sys.argv) == 1 else None)
