using ExcelDna.Integration;

namespace gmafffff.excel.udf.Excel.Сетка;

public interface ИМожетИзменятьВидимость
{
    XlCall.XlReturn Скрыть();
    XlCall.XlReturn Показать(double? размер = null, bool авторазмер = false);
    bool? ВидимЛи();
}