package hexcel;

import microsoft.office.interop.excel.Range;

abstract ExcelRange(Range) from Range to Range {
  // TODO runtime type check is required
  public inline function getNumber() : Float
    return this.Value2;

  public inline function setNumber(value : Float)
    this.Value2 = value;

  public inline function setFormula(formula : String)
    this.Formula = formula;

  public inline function clearPattern()
    setPattern(microsoft.office.interop.excel.XlPattern.xlPatternNone);

  public inline function setPattern(pattern : microsoft.office.interop.excel.XlPattern)
    this.Interior.Pattern = pattern;

  public inline function setThemeColor(color : microsoft.office.interop.excel.XlThemeColor)
    this.Interior.ThemeColor = color;

  public function map<TOut>(f : ExcelRange -> TOut) : Array<TOut> {
    var result = [],
        rows = this.Rows.Count,
        columns = this.Columns.Count;
    for(r in 0...rows)
      for(c in 0...columns)
        result.push(f(cell(r, c)));
    return result;
  }

  function cell(row : Int, column : Int) {
#if cs
    return this.get_Item(row+1, column+1);
#elseif js
    return this.Item(row+1, column+1);
#else
    return this[row+1][column+1];
#end
  }

#if cs
  @:functionCode('
    return ((object[,]) arr)[i, j];
  ')
  static function accessArray2(arr : Dynamic, i : Int, j : Int) : Dynamic {
    return null;
  }
#end
}