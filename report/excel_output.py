from openpyxl import Workbook
from dotenv import load_dotenv
import openpyxl
import random
import os
load_dotenv()

MIN_ROW=int(os.getenv('MIN_ROW'))
MAX_ROW=int(os.getenv('MAX_ROW'))
MIN_COL=int(os.getenv('MIN_COL'))
MAX_COL=int(os.getenv('MAX_COL'))
_SEP = ":"


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
    wb = Workbook()
    ws = wb.active
    for cell in cells_dict:
      print(cell,_SEP,cells_dict[cell])
      val = self.getValueFor(cells_dict[cell])
      ws[cell] = val
      print(f"setting {cell}={val}")

    wb.save(self.output_path)
  
  """
  extract value for the template cell
  """
  def getValueFor(self, value):
    return self.switch(value)
  

  """
  find the correct token and replace it with its new computed value
  """
  def switch(self,value)->any:
    print('value=>', value)
    match value:
      case "{o_t_1}":
        return self.computeResult()
      case "{o_t_2}":
        return self.computeResult()
      case "{o_t_3}":
        return self.computeResult()
      case _:
       return 0
      

  def computeResult(self):
    return random.randint(200, 4000)

