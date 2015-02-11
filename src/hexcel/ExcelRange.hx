package hexcel;

import microsoft.office.interop.excel.Range;

abstract ExcelRange(Range) from Range to Range {
  public inline function setNumber(value : Float)
    this.Value2 = value;
}