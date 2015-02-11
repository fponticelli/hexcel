package microsoft.office.interop.excel;

extern class Workbooks {
  public function Add(template : XlWBATemplate) : Dynamic {};
}

@:enum
abstract XlWBATemplate(Int)
{
  var xlWBATChart = 0xFFFFEFF3;
  var xlWBATExcel4IntlMacroSheet  = 0x4;
  var xlWBATExcel4MacroSheet  = 0x3;
  var xlWBATWorksheet  = 0xFFFFEFB9;
}