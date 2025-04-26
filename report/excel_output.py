from openpyxl import Workbook

"""
class to produce excel output
"""
class ExcelReportOutput:
  def __init__(self, template_path, output_path):
    self.template_path = template_path
    self.output_path = output_path
  """
  produce excel report output
  """
  def output(self):
    print(f"template:{self.template_path}")
    print(f"output:{self.output_path}")
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Origin CMM output'
    wb.save(self.output_path)