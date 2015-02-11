package microsoft.office.interop.excel;

extern class Range #if python implements ArrayAccess<Dynamic> #end {
  public var Value2 : Dynamic;
  public var Formula : Dynamic;
  var Cells(default,never) : Range;
  var Column(default,never) : Int;
  var Columns(default,never) : Range;
  var Count(default,never) : Int;
  var Row(default,never) : Int;
  var Rows(default,never) : Range;
  var Interior(default, never) : Interior;
#if js
  function Item(row : Int, col : Int) : Range;
#end
}