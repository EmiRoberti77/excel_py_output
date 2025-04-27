from openpyxl import Workbook
from dotenv import load_dotenv
import openpyxl
import os
load_dotenv()

MIN_ROW= int(os.getenv('MIN_ROW'))
MAX_ROW=int(os.getenv('MAX_ROW'))
MIN_COL=int(os.getenv('MIN_COL'))
MAX_COL=int(os.getenv('MAX_COL'))
"""
class to produce excel output
"""
class ExcelReportOutput:
  def __init__(self, template_path, output_path):
    self.template_path = template_path
    self.output_path = output_path
    print(self.template_path)
    print(self.output_path)
  
  """
  extract all values from the template
  """
  def read_template(self)->any:
    cells_dict = dict()
    try:
      template = openpyxl.load_workbook(self.template_path)
      print('template has been opened')
      ws = template.active
      for row in ws.iter_rows(min_row=MIN_ROW, max_row=MAX_ROW, min_col=MIN_COL, max_col=MAX_COL):
        for cell in row:
          if cell.value is not None:         
            cells_dict[cell.coordinate]=cell.value   
    except Exception as e: 
      print(e)
      print('failed to load file')
    return cells_dict


  """
  produce excel report output
  """
  def output(self):
    cells_dict = self.read_template()
    print(cells_dict)
    print(f"template:{self.template_path}")
    print(f"output:{self.output_path}")
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Origin CMM output'
    wb.save(self.output_path)