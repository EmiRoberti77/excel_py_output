from openpyxl import Workbook
import openpyxl

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
  def read_template(self):
    try:
      template = openpyxl.load_workbook(self.template_path)
      print('template has been opened')
    except Exception:
      print(Exception)
      print('failed to load file')




  """
  produce excel report output
  """
  def output(self):
    self.read_template()
    print(f"template:{self.template_path}")
    print(f"output:{self.output_path}")
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Origin CMM output'
    wb.save(self.output_path)