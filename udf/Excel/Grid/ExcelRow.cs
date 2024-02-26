using ExcelDna.Integration;

namespace gmafffff.excel.udf.Excel.Сетка;

public sealed record ExcelСтрока(int Номер, IntPtr ЛистИд) : РядСетки(Номер, ЛистИд) {
    private const double МинимальнаяВысота = .5;

    public double? Высота {
        get {
            var res = XlCall.TryExcel(XlCall.xlfGetCell, out var рез, 17, ПерваяЯчейка);

            if (res != XlCall.XlReturn.XlReturnSuccess || рез is not double высота) return null;
            return высота;
        }
    }

    public override ExcelReference ПерваяЯчейка
        => new(Номер, Номер, 0, 0, ЛистИд);

    public override bool? ВидимЛи() { return Высота >= МинимальнаяВысота; }

    // Приведение типов
    public static implicit operator ExcelReference(ExcelСтрока стр) {
        return new ExcelReference(стр.Номер, стр.Номер, 0, ExcelDnaUtil.ExcelLimits.MaxColumns - 1, стр.ЛистИд);
    }

    public static ExcelСтрока[] Преобразовать(ExcelReference ссылка) {
        return (from прямоугольник in ссылка.InnerReferences
                from номер in Enumerable.Range(прямоугольник.RowFirst,
                                               прямоугольник.RowLast - прямоугольник.RowFirst + 1)
                select new ExcelСтрока(номер, прямоугольник.SheetId))
           .ToArray();
    }

    public override ExcelReference GetAsExcelReference() { return this; }
}