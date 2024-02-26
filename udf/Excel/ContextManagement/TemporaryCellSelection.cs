using ExcelDna.Integration;

namespace gmafffff.excel.udf.Excel.УправлениеКонтекстом;

// Источник: https://excel-dna.net/docs/guides-advanced/async-macro-example-formatting-the-calling-cell-from-a-udf/
public sealed class ВременноеВыделениеЯчеек : XlCall, IDisposable {
    private readonly ОбновлениеЭкранаПриостановить _обновлениеЭкрана;
    private readonly object _стараяАктивнаяЯчейкаАктивногоЛиста;
    private readonly object _стараяАктивнаяЯчейкаЦелевогоЛиста;
    private readonly object _староеВыделениеАктивныйЛист;
    private readonly object _староеВыделениеЦелевогоЛиста;

    public ВременноеВыделениеЯчеек(ExcelReference временноеВыделение) {
        _обновлениеЭкрана = new ОбновлениеЭкранаПриостановить();

        // Запоминает старое выделение на активном листе
        _староеВыделениеАктивныйЛист        = Excel(xlfSelection);
        _стараяАктивнаяЯчейкаАктивногоЛиста = Excel(xlfActiveCell);

        // Переключение на лист, в котором нужно сделать выделение
        var целевойЛист = (string)Excel(xlSheetNm, временноеВыделение);
        Excel(xlcWorkbookSelect, целевойЛист);

        // Запоминает старое выделение на целевом листе
        _староеВыделениеЦелевогоЛиста      = Excel(xlfSelection);
        _стараяАктивнаяЯчейкаЦелевогоЛиста = Excel(xlfActiveCell);

        // Делает выделение ячеек
        Excel(xlcFormulaGoto, временноеВыделение);
    }

    public void Dispose() {
        Excel(xlcSelect, _староеВыделениеЦелевогоЛиста, _стараяАктивнаяЯчейкаЦелевогоЛиста);

        var старыйАктивныйЛист = (string)Excel(xlSheetNm, _староеВыделениеАктивныйЛист);
        Excel(xlcWorkbookSelect, старыйАктивныйЛист);

        Excel(xlcSelect, _староеВыделениеАктивныйЛист, _стараяАктивнаяЯчейкаАктивногоЛиста);

        _обновлениеЭкрана.Dispose();
    }
}