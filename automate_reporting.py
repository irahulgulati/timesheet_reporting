import xlrd as excel
import os
from datetime import date
import time
import sys
import requests
import json
import argparse


class timesheet:
    def __init__(self):
        self.parser = argparse.ArgumentParser(
            description='**Automated Timesheet Reporting**', usage="automate_reporting.py [-h] ||[filename.xlsx] [your name]")
        self.parser.add_argument('file', type=str,
                                 help='file name with .xlsx extension')
        self.parser.add_argument('name', type=str,
                                 help='Your Name')

        self.args = self.parser.parse_args()
        self._cachedTime = 0
        # Setting the base directory and file paths
        self.basedir = os.path.abspath(os.path.dirname(__file__))
        self.timesheetDir = self.basedir + '\\timesheets'
        self.slack_webhook = os.environ.get('Sankara_Webhook')

    # main logic

    def fileCheck(self):
        self.sheetlocation = (os.path.join(
            self.timesheetDir, self.args.file))

        # initiating the excel sheet
        self.wb = excel.open_workbook(self.sheetlocation)

        # setting up the start poing with cell index
        self.sheet = self.wb.sheet_by_index(0)
        self.sheet.cell_value(0, 0)
        self.numberofRows = self.sheet.nrows
        self.records = []
        for entry in range(1, self.numberofRows):
            self.records.append(self.sheet.row_values(entry))
        data = {
            "blocks": [
                {
                    "type": "divider"
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": "Name\n\n"+self.args.name
                    }
                },
                {
                    "type": "divider"
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": "Date: \n" + str(date.today()),
                    }
                },
                {
                    "type": "divider"
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": str(self.records[0])
                    }
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": str(self.records[1])
                    }
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": str(self.records[2])
                    }
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": str(self.records[3])
                    }
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": str(self.records[4])
                    }
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": str(self.records[5])
                    }
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": str(self.records[6])
                    }
                },
            ]
        }
        if os.stat(self.sheetlocation).st_mtime != self._cachedTime:
            requests.post(
                self.slack_webhook, data=json.dumps(data))
        else:
            # print(self._cachedTime)
            print("False")


if __name__ == "__main__":
    try:
        time.sleep(1)
        # instantiating the class object
        run = timesheet()
        run.fileCheck()
    except Exception as err:
        print(err)
        sys.exit(0)
