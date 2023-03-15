"""Powerpoint Merge
"""
from argparse import ArgumentParser
from glob import glob
from pathlib import Path
from typing import List

import win32com.client


def merge_ppts(output: Path, files: List[Path]):
    """Merge Powerpoint Presentations together

    Args:
        output (Path): Output path
        files (List[Path]): List of powerpoints to merge
    """
    ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
    new_prs = ppt_instance.Presentations.Add()
    new_prs.SaveAs(str(output.resolve()))

    for path in files:
        prs = ppt_instance.Presentations.Open(str(path.resolve()))
        prs.Slides.range(range(1, prs.Slides.Count + 1)).copy()
        ppt_instance.Presentations(output.name).Windows(1).Activate()
        new_prs.Application.CommandBars.ExecuteMSO('PasteSourceFormatting')
        prs.Close()

    new_prs.SaveAs(str(output.resolve()))
    new_prs.Close()
    ppt_instance.Quit()

def main() -> None:
    """Main function
    """
    parser = ArgumentParser()
    parser.add_argument('output', type=Path)
    parser.add_argument('files', nargs='+', type=str)

    args = parser.parse_args()
    arg_dict = vars(args)
    files = []
    for file_pattern in arg_dict['files']:
        files.extend(Path(file) for file in glob(file_pattern))
    arg_dict['files'] = files
    merge_ppts(**arg_dict)

if __name__ == "__main__":
    main()
