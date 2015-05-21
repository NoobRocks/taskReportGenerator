# taskReportGenerator
This script is used to generate the annual report required by my company policy. It groups the SVN log by a specified time delta and generates a report from the grouped logs.

## Requirements
* Lastest version of Python 2
* [python-docx](https://python-docx.readthedocs.org/en/latest/)

## Usage
1. Edit the file prefixed with "config_"
1. Put it and template document at the same directory as taskReportGenerator.py
1. cd to the directory and enter `python taskReportGenerator.py`