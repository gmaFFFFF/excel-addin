using ExcelDna.Integration;
using gmafffff.excel.udf.Excel.Константы;

namespace gmafffff.excel.udf.Excel.УправлениеКонтекстом;

public sealed class АвтоРасчетПриостановить : IDisposable
{
    private readonly bool _измененЛи;
    private bool _calcSave;
    private bool _date1904;
    private bool _iter;
    private double _maxChange;
    private int _maxNum;
    private bool _precision;
    private bool _saveValues;
    private НастройкиПересчета _typeNum;
    private bool _update;

    public АвтоРасчетПриостановить()
    {
        СохранитьТекущиеЗначения();
        if (_typeNum == НастройкиПересчета.Ручной)
            return;

        _измененЛи = true;
        XlCall.Excel(XlCall.xlcOptionsCalculation,
            (int)НастройкиПересчета.Ручной,
            _iter, _maxNum, _maxChange, _update, _precision, _date1904, _calcSave, _saveValues);
    }

    public void Dispose()
    {
        ВосстановитьСохраненныеЗначения();
    }

    private void СохранитьТекущиеЗначения()
    {
        _typeNum = (НастройкиПересчета)Convert.ToInt32(XlCall.Excel(XlCall.xlfGetDocument, 14));
        _iter = (bool)XlCall.Excel(XlCall.xlfGetDocument, 15);
        _maxNum = Convert.ToInt32(XlCall.Excel(XlCall.xlfGetDocument, 16));
        _maxChange = (double)XlCall.Excel(XlCall.xlfGetDocument, 17);
        _update = (bool)XlCall.Excel(XlCall.xlfGetDocument, 18);
        _precision = (bool)XlCall.Excel(XlCall.xlfGetDocument, 19);
        _date1904 = (bool)XlCall.Excel(XlCall.xlfGetDocument, 20);
        _calcSave = (bool)XlCall.Excel(XlCall.xlfGetDocument, 33);
        _saveValues = (bool)XlCall.Excel(XlCall.xlfGetDocument, 43);
    }

    private XlCall.XlReturn ВосстановитьСохраненныеЗначения()
    {
        return !_измененЛи
            ? XlCall.XlReturn.XlReturnSuccess
            : XlCall.TryExcel(XlCall.xlcOptionsCalculation, out _,
                (int)_typeNum,
                _iter, _maxNum, _maxChange, _update, _precision, _date1904, _calcSave, _saveValues);
    }
}