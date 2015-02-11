import hexcel.*;

class Main {
  public static function main() {
    var excel = Excel.create(),
      wb = excel.addWorkbook(),
      ws = wb.get(0),
      range = ws.cell("A1");

    range.setNumber(6);
    ws.cellAt(0, 1).setNumber(10);
    ws.cellAt(2, 3).setNumber(666);
  }
}