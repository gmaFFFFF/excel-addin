using ExcelDna.Integration;

namespace gmafffff.excel.udf.AddIn;

public record ExcelСтрока(int Номер, IntPtr ЛистИд) : ИExcelЭлемВидимость {
    private const double МинимальнаяВысота = .5;

    public bool? ВидимЛи() {
        var res = XlCall.TryExcel(XlCall.xlfGetCell, out var рез, 17, (ExcelReference)this);

        if (res != XlCall.XlReturn.XlReturnSuccess || рез is not double высота)
            return null;
        return !(высота <= МинимальнаяВысота);
    }

    public ExcelReference AsExcelReference() {
        return this;
    }

    public XlCall.XlReturn Скрыть() {
        return ((ИExcelЭлемВидимость)this).ВидимостьСтрСтлб(XlCall.xlcRowHeight, -1);
    }

    public XlCall.XlReturn Показать(double? высота = null) {
        return ((ИExcelЭлемВидимость)this).ВидимостьСтрСтлб(XlCall.xlcRowHeight, высота);
    }

    public static implicit operator ExcelReference(ExcelСтрока стр) {
        return new ExcelReference(стр.Номер, стр.Номер, 0, 0, стр.ЛистИд);
    }

    public static ExcelСтрока[] Преобразовать(ExcelReference ссылка) {
        return (from прямоуг in ссылка.InnerReferences
                from номер in Enumerable.Range(прямоуг.RowFirst, прямоуг.RowLast - прямоуг.RowFirst + 1)
                select new ExcelСтрока(номер, прямоуг.SheetId))
            .ToArray();
    }
}