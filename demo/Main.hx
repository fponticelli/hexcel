import hexcel.*;

using thx.Arrays;

class Main {
  public static function main() {
    var excel = Excel.create(),
        wb = excel.addWorkbook(),
        ws = wb.get(0),
        values = [99.99, 1.11, 4.25, 98.83, 11.48, 14.58, 23.85, 85.33, 61.92, 53.74];

    values.plucki(ws.cellAt(i, 0).setNumber(_));
    ws.cellAt(values.length + 1, 0).setFormula('=VAR.P(A1:A${values.length})');

    var range = ws.rangeAt(0, 0, values.length - 1, 0);
    markLow(range, 10);

    makeChart(ws, range);
  }

  static function markLow(range : ExcelRange, value : Float) {
    range.map(function(cell) {
      if(cell.getNumber() <= value) {
        cell.setPattern(microsoft.office.interop.excel.XlPattern.xlPatternSolid);
        cell.setThemeColor(microsoft.office.interop.excel.XlThemeColor.xlThemeColorAccent6);
      }
    });
  }

  static function makeChart(sheet : ExcelWorksheet, range : ExcelRange) {
    sheet.addChart(xlColumnClustered, 120, 20, 500, 500, range);
  }
}