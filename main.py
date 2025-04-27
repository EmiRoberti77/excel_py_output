from report.excel_output import ExcelReportOutput
from dotenv import load_dotenv
import os
load_dotenv()
TEMPLATE = os.getenv("TEMPLATE_INPUT_PATH")
OUTPUT = os.getenv("REPORT_OUTPUT_PATH")
INPUT = os.getenv("REPORT_INPUT_PATH")

excel_output = ExcelReportOutput(TEMPLATE, OUTPUT, INPUT)
excel_output.output()




 
