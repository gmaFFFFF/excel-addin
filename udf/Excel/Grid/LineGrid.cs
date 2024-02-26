using ExcelDna.Integration;
using gmafffff.excel.udf.Excel.Константы;

namespace gmafffff.excel.udf.Excel.Сетка;

/// <summary>
///     Строка или столбец
/// </summary>
public abstract record РядСетки(int Номер, IntPtr ЛистИд) : ИМожетИзменятьВидимость {
    public abstract ExcelReference ПерваяЯчейка { get; }
    public abstract ExcelReference GetAsExcelReference();

    # region Видимость ряда

    public static XlCall.XlReturn Скрыть(ExcelReference reference, Type тип) {
        var функцияXlCallДляРазмераРяда = ФункцияXlCallДляРазмераРяда(тип);
        return XlCall.TryExcel(функцияXlCallДляРазмераРяда, out _, null, reference, false,
                               (int)РядСеткиВидимость.Скрыть);
    }

    public static XlCall.XlReturn Показать(ExcelReference reference, Type тип, double размер) {
        var функцияXlCallДляРазмераРяда = ФункцияXlCallДляРазмераРяда(тип);
        return XlCall.TryExcel(функцияXlCallДляРазмераРяда, out _, размер, reference, false);
    }

    public static XlCall.XlReturn ПоказатьКакБыло(ExcelReference reference, Type тип) {
        var функцияXlCallДляРазмераРяда = ФункцияXlCallДляРазмераРяда(тип);
        return XlCall.TryExcel(функцияXlCallДляРазмераРяда, out _, null, reference, false,
                               (int)РядСеткиВидимость.Восстановить);
    }

    public static XlCall.XlReturn ПоказатьАвто(ExcelReference reference, Type тип) {
        var функцияXlCallДляРазмераРяда = ФункцияXlCallДляРазмераРяда(тип);
        return XlCall.TryExcel(функцияXlCallДляРазмераРяда, out _, null, reference, false, (int)РядСеткиВидимость.Авто);
    }

    public static XlCall.XlReturn ПоказатьПоУмолчанию(ExcelReference reference, Type тип) {
        var функцияXlCallДляРазмераРяда = ФункцияXlCallДляРазмераРяда(тип);
        return XlCall.TryExcel(функцияXlCallДляРазмераРяда, out _, null, reference, true);
    }

    /// <summary>
    ///     Метод скрытия/отображения ряда Excel(строка/столбец) через установку размера (высота/ширина)
    /// </summary>
    /// <param name="reference">
    ///     Ссылка на элемент сетки.
    /// </param>
    /// <param name="тип">
    ///     Тип элемента сетки, который нужно скрыть показать
    ///     <list type="bullet">
    ///         <listheader>
    ///             <term>Значение</term>
    ///             <description>Описание</description>
    ///         </listheader>
    ///         <item>
    ///             <term>typeof(ExcelСтрока)</term>
    ///             <description>Строка</description>
    ///         </item>
    ///         <item>
    ///             <term>typeof(ExcelСтолбец)</term>
    ///             <description>Столбец</description>
    ///         </item>
    ///     </list>
    /// </param>
    /// <param name="размер">
    ///     необязательный размер:
    ///     если ≤0, то элемент скрывается;
    ///     если null или >0, то элемент отображается.
    /// </param>
    /// <param name="авторазмер">
    ///     Позволить Excel самому определить размер ряда.
    /// </param>
    /// <returns></returns>
    public static XlCall.XlReturn ИзменитьВидимость(ExcelReference reference, Type тип,
                                                    double? размер = null, bool авторазмер = false) {
        return (размер, авторазмер) switch {
            // Скрыть
            (<= 0, _) => Скрыть(reference, тип),

            // Показать
            (_, true)  => ПоказатьАвто(reference, тип),
            ({ } р, _) => Показать(reference, тип, р),
            (null, _)  => ПоказатьКакБыло(reference, тип)
            //_ => РядСетки.ПоказатьПоУмолчанию(reference, функцияXlCallДляРазмераРяда)
        };
    }

    public static void ИзменитьВидимость(IEnumerable<ExcelReference> references, Type тип,
                                         double? размер = null, bool авторазмер = false) {
        var refПоЛисту = from r in references
                         group r by r.SheetId;

        foreach (var лист in refПоЛисту) ИзменитьВидимость(new ExcelReference(лист), тип, размер, авторазмер);
    }

    protected XlCall.XlReturn ИзменитьВидимость(double? размер = null, bool авторазмер = false) {
        if (ВидимЛи() is true ^ размер is <= 0) return XlCall.XlReturn.XlReturnAbort;
        return ИзменитьВидимость(GetAsExcelReference(), GetType(), размер, авторазмер);
    }

    public XlCall.XlReturn Скрыть() { return Скрыть(GetAsExcelReference(), GetType()); }

    public XlCall.XlReturn Показать(double? размер = null, bool авторазмер = false) {
        return ИзменитьВидимость(размер, авторазмер);
    }

    public abstract bool? ВидимЛи();

    private static int ФункцияXlCallДляРазмераРяда(Type тип) {
        if (тип == typeof(ExcelСтрока)) return XlCall.xlcRowHeight;
        if (тип == typeof(ExcelСтолбец)) return XlCall.xlcColumnWidth;

        throw new NotSupportedException("Не поддерживаемый элемент ряда сетки");
    }

    # endregion
}