import sys
import click
from dargalpy.dargalpy import run_process


@click.group()
@click.version_option("1.0.0")
def main():
    print("Time & Expenses auto dargal report...")
    pass
@main.command()
@click.option(
    "--filepath", 
    type=click.STRING,
    required=True,
    help="Full path to folder path"
)
def run(filepath):
    run_process(filepath)

if __name__ == '__main__':
    args = sys.argv
    if "--help" in args or len(args) == 1:
        print("Time & Expenses auto dargal report...")
    main()