'''E4E Weekly Meeting Powerpoint Tool
'''
import argparse
from pathlib import Path
from random import shuffle
from shutil import copy
from tempfile import TemporaryDirectory
from typing import List

import yaml

from e4e_pptx_merge.pptx_merge import merge_ppts


class E4eMeetingMerge:
    """E4E Meeting Merge Tool
    """
    # pylint: disable=too-many-instance-attributes,too-few-public-methods
    # For simple application
    def __init__(self):
        self.parser = argparse.ArgumentParser()
        self.parser.add_argument('--job_file', '-j', type=Path, required=True)
        self.parser.add_argument('--week', '-w', type=int, required=True)
        self.parser.add_argument('--output', '-o', type=Path, default=Path('.'))
        self.parser.add_argument('--order', choices=['sequence', 'random'], type=str,
                                 default='sequence')

        args = self.parser.parse_args()

        self.job_file_path: Path = args.job_file
        self.week_num: int = args.week
        self.output_dir: Path = args.output
        self.order: str = args.order

        with open(self.job_file_path, 'r', encoding='ascii') as handle:
            job_file_params = yaml.safe_load(handle)
        self.presentation_paths = {str(project): Path(path)
                              for project, path in job_file_params['paths'].items()}
        self.this_weeks_projects: List[str] = job_file_params['schedule'][self.week_num]
        self.quarter: str = job_file_params['quarter']
        self.general_ppt_path: Path = Path(job_file_params['general'])

    def run(self):
        """Application entry point
        """
        presentation_paths = {project:path.joinpath(self.quarter)
                              for project, path in self.presentation_paths.items()
                              if path.joinpath(self.quarter).exists()
                              and project in self.this_weeks_projects}

        if self.order == 'random':
            shuffle(self.this_weeks_projects)
        print(self.this_weeks_projects)
        with TemporaryDirectory() as tempdir:
            merge_output = Path(tempdir).joinpath('output.pptx')
            starter = self.general_ppt_path.joinpath(self.quarter, f'Week {self.week_num}.pptx')
            presentations_to_merge: List[Path] = [starter]
            for project in self.this_weeks_projects:
                presentation_dir = presentation_paths[project]
                presentation = presentation_dir.joinpath(f'Week {self.week_num}.pptx')
                if presentation.exists():
                    presentations_to_merge.append(presentation)
            print(presentations_to_merge)
            merge_ppts(merge_output, presentations_to_merge)
            copy(merge_output, self.output_dir.joinpath(f'Week {self.week_num} final.pptx'))

def main():
    """Main app logic
    """
    E4eMeetingMerge().run()

if __name__ == '__main__':
    main()
