using ExcelDna.Integration;

namespace gmafffff.excel.udf.Excel.Сетка;

public sealed record ExcelСтолбец(int Номер, IntPtr ЛистИд) : РядСетки(Номер, ЛистИд)
{
    private const double МинимальнаяШирина = .05;

    public double? Ширина
    {
        get
        {
            var res1 = XlCall.TryExcel(XlCall.xlfGetCell, out var резПраво, 44, ПерваяЯчейка);
            var res2 = XlCall.TryExcel(XlCall.xlfGetCell, out var резЛево, 42, ПерваяЯчейка);

            if (res1 != XlCall.XlReturn.XlReturnSuccess || res2 != XlCall.XlReturn.XlReturnSuccess
                                                        || резПраво is not double право || резЛево is not double лево)
                return null;
            return право - лево;
        }
    }

    public override ExcelReference ПерваяЯчейка
        => new(0, 0, Номер, Номер, ЛистИд);

    public override bool? ВидимЛи()
    {
        return Ширина >= МинимальнаяШирина;
    }

    // Приведение типов
    public static implicit operator ExcelReference(ExcelСтолбец стлб)
    {
        return new ExcelReference(0, ExcelDnaUtil.ExcelLimits.MaxRows - 1, стлб.Номер, стлб.Номер, стлб.ЛистИд);
    }

    public static ExcelСтолбец[] Преобразовать(ExcelReference ссылка)
    {
        return (from прямоугольник in ссылка.InnerReferences
                from номер in Enumerable.Range(прямоугольник.ColumnFirst,
                    прямоугольник.ColumnLast - прямоугольник.ColumnFirst + 1)
                select new ExcelСтолбец(номер, прямоугольник.SheetId))
            .ToArray();
    }

    public override ExcelReference GetAsExcelReference()
    {
        return this;
    }
}