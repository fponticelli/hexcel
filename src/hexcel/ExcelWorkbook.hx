package hexcel;

import microsoft.office.interop.excel.Workbook;

abstract ExcelWorkbook(Workbook) from Workbook to Workbook {
  public inline function get(index : Int) : ExcelWorksheet {
#if cs
    return this.Worksheets.get_Item(index + 1);
#else
    return this.Worksheets.Item(index + 1);
#end
  }
}