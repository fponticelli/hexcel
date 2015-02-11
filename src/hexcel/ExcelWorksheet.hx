package hexcel;

import microsoft.office.interop.excel.Worksheet;

abstract ExcelWorksheet(Worksheet) from Worksheet to Worksheet {
  public inline function cell(ref : String) : ExcelRange
#if cs
    return this.get_Range(ref, ref);
#else
    return this.Range(ref, ref);
#end

  public inline function cellAt(row : Int, col : Int) : ExcelRange
    return cell(posToRef(row, col));

  static var alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  public static function posToRef(row : Int, col : Int) {
    var c = "";
    while (col > alpha.length) {
      c += Math.floor(col / alpha.length);
    }
    c += alpha.substr(col, 1);
    return '$c${row+1}';
  }
}