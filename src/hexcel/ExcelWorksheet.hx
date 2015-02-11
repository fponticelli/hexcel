package hexcel;

import microsoft.office.interop.excel.Worksheet;
import microsoft.office.interop.excel.XlChartType;

abstract ExcelWorksheet(Worksheet) from Worksheet to Worksheet {
  public inline function cell(ref : String) : ExcelRange
    return range(ref, ref);

  public inline function cellAt(row : Int, col : Int) : ExcelRange
    return cell(posToRef(row, col));

  public inline function range(ref1 : String, ref2 : String) : ExcelRange
#if cs
    return this.get_Range(ref1, ref2);
#else
    return this.Range(ref1, ref2);
#end

  public inline function rangeAt(row1 : Int, col1 : Int, row2 : Int, col2 : Int) : ExcelRange
    return range(posToRef(row1, col1), posToRef(row2, col2));

  static var alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  public static function posToRef(row : Int, col : Int) {
    var c = "";
    while (col > alpha.length) {
      c += Math.floor(col / alpha.length);
    }
    c += alpha.substr(col, 1);
    return '$c${row+1}';
  }

  public function addChart(type : XlChartType, x : Int, y : Int, width : Int, height : Int, range : ExcelRange) : ExcelChart {
    var chart = this.Shapes.AddChart(type, x, y, width, height).Chart;
    chart.SetSourceData(range, microsoft.office.interop.excel.XlRowCol.xlColumns);
    return chart;
  }
}