package hexcel;

import microsoft.office.interop.excel.Application;
import microsoft.office.interop.excel.Workbook;
import microsoft.office.interop.excel.Workbooks;

abstract Excel(Application) from Application to Application {
  public static function create(visible = true) : Excel {
    var excel = getNative();
    excel.Visible = visible;
    return excel;
  }

  public inline function addWorkbook() : ExcelWorkbook {
    return (this.Workbooks.Add(XlWBATemplate.xlWBATWorksheet) : ExcelWorkbook);
  }

  static function getNative() : Application {
#if python
    return Win32.gencache.EnsureDispatch('Excel.Application');
#elseif js
    if(untyped __js__("typeof window !== 'undefined' && 'ActiveXObject' in window"))
      return untyped __js__("new ActiveXObject")("Excel.Application");
    if(untyped __js__("typeof require !== 'undefined'"))
      return untyped __js__("require")("win32ole").client.Dispatch("Excel.Application");
    return throw "The JS Host doesn't support a mean to communicate with Excel";
#elseif cs
    return new microsoft.office.interop.excel.ApplicationClass();
#else
    return throw "Target platform is not implemented";
#end
  }
}

#if python
@:pythonImport("win32com.client")
extern class Win32 {
  public static var gencache : {
    EnsureDispatch : String -> Dynamic
  };
}
#end