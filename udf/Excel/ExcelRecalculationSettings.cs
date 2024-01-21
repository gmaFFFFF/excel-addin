using ExcelDna.Integration;

namespace gmafffff.excel.udf.AddIn;

public class ExcelПересчетНастройки {
    public enum ExcelПересчет {
        Авто = 1,
        АвтоБезТабл = 2,
        Ручной = 3
    }

    private bool сохраненЛи;
    private bool calc_save;
    private bool date_1904;
    private bool iter;
    private double max_change;
    private int max_num;
    private bool precision;
    private bool save_values;
    private ExcelПересчет type_num;
    private bool update;

    private void СохранитьТекущиеЗначения() {
        type_num = (ExcelПересчет)Convert.ToInt32(XlCall.Excel(XlCall.xlfGetDocument, 14));
        iter = (bool)XlCall.Excel(XlCall.xlfGetDocument, 15);
        max_num = Convert.ToInt32(XlCall.Excel(XlCall.xlfGetDocument, 16));
        max_change = (double)XlCall.Excel(XlCall.xlfGetDocument, 17);
        update = (bool)XlCall.Excel(XlCall.xlfGetDocument, 18);
        precision = (bool)XlCall.Excel(XlCall.xlfGetDocument, 19);
        date_1904 = (bool)XlCall.Excel(XlCall.xlfGetDocument, 20);
        calc_save = (bool)XlCall.Excel(XlCall.xlfGetDocument, 33);
        save_values = (bool)XlCall.Excel(XlCall.xlfGetDocument, 43);
        сохраненЛи = true;
    }

    public XlCall.XlReturn ВосстановитьСохраненныеЗначения() {
        return XlCall.TryExcel(XlCall.xlcOptionsCalculation, out _,
            (int)type_num,
            iter, max_num, max_change, update, precision, date_1904, calc_save, save_values);
    }

    public XlCall.XlReturn ОтключитьАвтоРасчет() {
        СохранитьТекущиеЗначения();
        if (type_num == ExcelПересчет.Ручной)
            return XlCall.XlReturn.XlReturnSuccess;

        return XlCall.TryExcel(XlCall.xlcOptionsCalculation, out _,
            (int)ExcelПересчет.Ручной,
            iter, max_num, max_change, update, precision, date_1904, calc_save, save_values);
    }
}