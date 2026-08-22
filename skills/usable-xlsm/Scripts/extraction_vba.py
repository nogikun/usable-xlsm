from pathlib import Path
import sys

import typer

from usable_xlsm import extract_vba


# Excel's dialogs and VBA errors are localised, so force UTF-8 on the streams:
# otherwise Windows encodes them in the console codepage and any tool reading
# this output gets mojibake instead of the error message.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


app = typer.Typer(add_completion=False)


@app.command()
def extract(
    input_path: Path = typer.Option(
        ...,
        "--input",
        "-i",
        exists=True,
        file_okay=True,
        dir_okay=False,
        help="Path to the XLSM workbook.",
    ),
    output: Path | None = typer.Option(
        None,
        "--output",
        "-o",
        file_okay=False,
        help="Directory where VBA files are saved.",
    ),
) -> None:
    modules = extract_vba(input_path)

    if output is None:
        for name, code in modules.items():
            typer.echo(f"--- {name} ---")
            typer.echo(code, nl=not code.endswith("\n"))
        return

    output.mkdir(parents=True, exist_ok=True)
    for name, code in modules.items():
        (output / Path(name).name).write_text(code, encoding="utf-8", newline="")


if __name__ == "__main__":
    app(args=["--help"] if len(sys.argv) == 1 else None)
