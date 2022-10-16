using ExcelDna.Integration;

namespace gmafffff.excel.udf.AddIn;

public static class Функции {
    [ExcelFunction(Name = "РублиПрописью", Category = "Функции от gmaFFFFF",
        Description = "Отображает сумму в рублях прописью")]
    public static string РублиПрописью(
        [ExcelArgument(Name = "суммаРублей", Description = "сумма, которую необходимо написать прописью")]
        double сумма,
        [ExcelArgument(Name = "формат",
            Description =
                @"ч[n](сумма, n - знаков после запятой), б/д[n][т[з]] (целая/дробная часть, т - текстом, з - с заглавной, n - ширина), р/к[с] (валюта базовая/дробная, с - сокращенная). Пример: ""ч2 рс (бтз р д2 к)""")]
        string формат = "") {
        return ДеньгиПрописью.ДеньгиПрописью.РублиПрописью(сумма, формат);
    }
}