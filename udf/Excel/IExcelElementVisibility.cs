using ExcelDna.Integration;

namespace gmafffff.excel.udf.AddIn;

public interface ИExcelЭлемВидимость {
    public enum ExcelСтрСтлбВидимость {
        Скрыть = 1,
        Восстановить = 2,
        Авто = 3
    }

    XlCall.XlReturn Скрыть();
    XlCall.XlReturn Показать(double? размер = null);
    bool? ВидимЛи();
    ExcelReference AsExcelReference();

    /// <summary>
    ///     Метод скрытия элемента (строка/столбец) через установку размера (высота/ширина)
    /// </summary>
    /// <param name="размер">необязательный размер. Если ≤0, то элемент скрывается</param>
    /// <returns></returns>
    protected internal XlCall.XlReturn ВидимостьСтрСтлб(int excelFunc, double? размер = null) {
        return (ВидимЛи(), размер) switch {
            // Скрыть
            (false, <= 0) => XlCall.XlReturn.XlReturnAbort,
            (_, <= 0) => XlCall.TryExcel(excelFunc, out _,
                null, AsExcelReference(), false, (int)ExcelСтрСтлбВидимость.Скрыть),
            // Показать
            (true, _) => XlCall.XlReturn.XlReturnAbort,
            (false, { } в) => XlCall.TryExcel(excelFunc, out _,
                в, AsExcelReference(), false),
            (false, null) => XlCall.TryExcel(excelFunc, out _,
                null, AsExcelReference(), false, (int)ExcelСтрСтлбВидимость.Восстановить),
            _ => XlCall.TryExcel(excelFunc, out _,
                null, AsExcelReference(), false, (int)ExcelСтрСтлбВидимость.Авто)
        };
    }
}