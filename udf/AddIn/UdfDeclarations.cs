using ExcelDna.Integration;

namespace gmafffff.excel.udf.AddIn;

public static class Функции {
    private const string МояКатегория = "Функции от gmaFFFFF";

    [ExcelFunction(Name = "РублиПрописью", Category = МояКатегория,
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

    [ExcelFunction(Name = "ОкруглГаус", Category = МояКатегория,
        Description = "Округление по Гауссу до ближайшего четного знака")]
    public static double ОкруглГ(
        [ExcelArgument(Name = "число", Description = "округляемое число")]
        double число,
        [ExcelArgument(Name = "знаков",
            Description = @"число знаков, до которых происходит округление.
Если число отрицательное, то округление до десятков перед запятой.
Максимум 15 знаков")]
        short знаков) {
        return (знаков, Math.Pow(10, -знаков)) switch {
            (> 15, _) => Math.Round(число, 15, MidpointRounding.ToEven),
            (>= 0, _) => Math.Round(число, знаков, MidpointRounding.ToEven),
            (< 0, var степень) => Math.Round(число / степень, 0, MidpointRounding.ToEven) * степень
        };
    }
}