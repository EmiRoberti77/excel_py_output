from report.excel_output import ExcelReportOutput
from dotenv import load_dotenv
import os
import sys
load_dotenv()
TEMPLATE_1 = os.getenv("TEMPLATE_INPUT_PATH_1")
TEMPLATE_2 = os.getenv("TEMPLATE_INPUT_PATH_2")
OUTPUT = os.getenv("REPORT_OUTPUT_PATH")
INPUT = os.getenv("REPORT_INPUT_PATH")

program = sys.argv[0]
template = sys.argv[1]
if template is None:
  print("missing template option")
else:
  if template == "1":
    template = TEMPLATE_1
  else: 
    template = TEMPLATE_2

  excel_output = ExcelReportOutput(template, OUTPUT, INPUT)
excel_output.output()




 
