namespace gmafffff.excel.udf.AddIn;

public interface ИExcelКоманда { }

public record ExcelКомандаВидимости(ИExcelЭлемВидимость элем, bool показать, double? размер = null) : ИExcelКоманда;