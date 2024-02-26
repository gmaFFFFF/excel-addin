using ExcelDna.Integration;

namespace gmafffff.excel.udf.Excel.УправлениеКонтекстом;

public sealed class ОбновлениеЭкранаПриостановить : XlCall, IDisposable {
    private readonly bool _измененЛи;
    private readonly bool _обновлятьЭкран;

    public ОбновлениеЭкранаПриостановить() {
        _обновлятьЭкран = (bool)Excel(xlfGetWorkspace, 40);
        if (!_обновлятьЭкран) return;
        _измененЛи = true;
        Excel(xlcEcho, false);
    }

    public void Dispose() {
        if (_измененЛи) Excel(xlcEcho, _обновлятьЭкран);
    }
}