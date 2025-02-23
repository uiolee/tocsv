import pandas as pd
from pathlib import Path


def to_csv(inputPath: Path, output_dir: str = "tocsv", output_encoding='utf-8-sig', ):
    def get_output_path(input_Path: Path, output_dir: str, sheet_name: str):
        parent = Path(input_Path.parent).joinpath(output_dir)
        parent.mkdir(parents=True, exist_ok=True)
        target = parent.joinpath(f"{input_Path.stem}-{sheet_name}{input_Path.suffix}.csv")
        return target

    with pd.ExcelFile(inputPath, engine="calamine") as file:
        print()
        print(f"processing: {inputPath}")
        sheets = pd.read_excel(file, sheet_name=None, dtype=str)
        for sheet_name, sheet_df in sheets.items():
            output_csv = get_output_path(inputPath, output_dir, sheet_name)
            sheet_df.to_csv(output_csv, index=False, encoding=output_encoding, )
            print(f"    wrote: {output_csv}")


if __name__ == '__main__':
    inputPaths = [Path('sheets.ods'), Path('sheets.xlsx'), Path('sheets.xls')]
    output_dir = 'dir1'

    for inputPath in inputPaths:
        to_csv(inputPath, output_dir)
