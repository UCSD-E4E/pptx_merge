"""Example setup file
"""
from setuptools import setup, find_packages

setup(
    name='e4e_pptx_merge',
    version='0.1.0.0',
    author='UCSD Engineers for Exploration',
    author_email='e4e@eng.ucsd.edu',
    entry_points={
        'console_scripts': [
            'pptx_merge = e4e_pptx_merge.pptx_merge:main',
            'e4e_meeting_tool = e4e_pptx_merge.e4e_pptx_merge:main',
        ]
    },
    packages=find_packages(),
    install_requires=[
    'pywin32',
    'pyyaml',
    'pytest',
    ],
    extras_require={
        'dev': [
            'pytest',
            'coverage',
            'pylint',
            'wheel',
        ]
    },
)
