import typer
from typing import List
import csv
from stock_cli.stocks_table import StocksTable

app = typer.Typer()


@app.command()
def get(
    stocks: List[str],
    custom_name: bool = typer.Option(
        False,
        "-n",
        "--name",
        help="Specify a custom name."
        "You will be prompted  regardless when loading from a file.",
    ),
    from_file: bool = typer.Option(
        False, "-f", "--file", help="Import stocks from a file"
    ),
    period: str = typer.Option(
        "1y",
        "-p",
        "--period",
        help="Period of time data will be downloaded for. "
        "\n Valid periods: #d, #w, #m, #y, max",
        show_default=True,
    ),
    interval: str = typer.Option(
        "1d",
        "-i",
        "--interval",
        help="Valid intervals: "
        "1m,2m,5m,15m,30m,60m,90m,1h,1d,5d,1wk,1mo,3mo"
        "Intraday data cannot extend last 60 days",
        show_default=True,
    ),
    end: str = typer.Option(
        "Today",
        "-e",
        "--end",
        help="Download end date string (YYYY-MM-DD)",
        show_default=True,
    ),
):
    """
    Download stock data from yahoo finance, calculate technical indicators and export
    result to excel file.
    """
    if from_file:
        stocks = stocks[0]
        with open(stocks) as file:
            reader = csv.reader(file)
            data = list(reader)
            data = sorted([item[0] for item in data])
            name = typer.prompt("Choose a name for the excel sheet")
    else:
        data = sorted(stocks)
        if custom_name:
            name = typer.prompt("Choose a name for the excel sheet")
        else:
            name = " ".join(data)

    typer.echo("Downloading Data...")
    stocks_table = StocksTable(data, period, interval, end)
    stocks_table.export_excel(name)
    typer.echo("A excel file was successfully created!")
