using ExcelDna.Integration;

namespace gmafffff.excel.udf.AddIn;

public record ExcelСтолбец(int Номер, IntPtr ЛистИд) : ИExcelЭлемВидимость {
    private const double МинимальнаяШирина = .05;

    public bool? ВидимЛи() {
        var res1 = XlCall.TryExcel(XlCall.xlfGetCell, out var рез_право, 44, (ExcelReference)this);
        var res2 = XlCall.TryExcel(XlCall.xlfGetCell, out var рез_лево, 42, (ExcelReference)this);

        if (res1 != XlCall.XlReturn.XlReturnSuccess
            || res2 != XlCall.XlReturn.XlReturnSuccess
            || рез_право is not double право
            || рез_лево is not double лево)
            return null;
        var ширина = право - лево;
        return !(ширина <= МинимальнаяШирина);
    }

    public ExcelReference AsExcelReference() {
        return this;
    }

    public XlCall.XlReturn Скрыть() {
        return ((ИExcelЭлемВидимость)this).ВидимостьСтрСтлб(XlCall.xlcColumnWidth, -1);
    }

    public XlCall.XlReturn Показать(double? высота = null) {
        return ((ИExcelЭлемВидимость)this).ВидимостьСтрСтлб(XlCall.xlcColumnWidth, высота);
    }

    public static implicit operator ExcelReference(ExcelСтолбец стлб) {
        return new ExcelReference(0, 0, стлб.Номер, стлб.Номер, стлб.ЛистИд);
    }

    public static ExcelСтолбец[] Преобразовать(ExcelReference ссылка) {
        return (from прямоуг in ссылка.InnerReferences
                from номер in Enumerable.Range(прямоуг.ColumnFirst, прямоуг.ColumnLast - прямоуг.ColumnFirst + 1)
                select new ExcelСтолбец(номер, прямоуг.SheetId))
            .ToArray();
    }
}