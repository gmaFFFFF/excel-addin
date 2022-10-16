using ExcelDna.Integration;
using ExcelDna.Registration;

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

    #region Http

    [ExcelAsyncFunction(Name = "HttpGet", Category = МояКатегория,
        Description = @$"Выполняет Get запрос по адресу и возвращает json объект с полями:
{nameof(HttpКлиент.HttpКлиент.ОтветHttp.ДатаЗапроса)}; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Статус)}; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Содержимое)}; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Заголовки)} — заголовк ответа; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Заголовки2)} — заголовк содержимого; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Куки)} — куки.
Не забывайте про ограничения по количеству символов в ячейке")]
    public static async Task<object> HttpGet(
        [ExcelArgument(Name = "адрес", Description = "Url адрес, по которому необходимо сделать запрос")]
        string адрес,
        [ExcelArgument(Name = "jsonPath", Description =
            @"Задаётся в формате JSONPath и позволяет сразу перейти к нужному элементу Json.
Подробнее о формате JSONPath по ссылке:
https://danielaparker.github.io/JsonCons.Net/articles/JsonPath/JsonConsJsonPath.html")]
        string jsonPath = "$",
        [ExcelArgument(Name = "заголовки", Description =
            "Диапазон: столбец с названиями и столбцы со значениями заголовка запроса (необязательно)")]
        string[,]? заголовки = null,
        CancellationToken ct = default) {
        // Предотвращает выполнение пока запущен мастер функций
        if (ExcelDnaUtil.IsInFunctionWizard()) return "";

        if (заголовки?.Length == 1)
            заголовки = null;

        var ответ = await HttpКлиент.HttpКлиент.GetАсинх(адрес, заголовки, ct).ConfigureAwait(false);

        if (ответ is null)
            return ExcelError.ExcelErrorNA;

        if (string.IsNullOrWhiteSpace(jsonPath) || jsonPath == "$")
            return ответ;

        return JsonКлиент.JsonКлиент.JsonPathНайди(ответ, jsonPath);
    }

    [ExcelAsyncFunction(Name = "HttpPost", Category = МояКатегория,
        Description = $@"Выполняет Post запрос по адресу и возвращает json объект с полями:
{nameof(HttpКлиент.HttpКлиент.ОтветHttp.ДатаЗапроса)}; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Статус)}; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Содержимое)}; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Заголовки)} — заголовк ответа; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Заголовки2)} — заголовк содержимого; {nameof(HttpКлиент.HttpКлиент.ОтветHttp.Куки)} — куки
Не забывайте про ограничения по количеству символов в ячейке")]
    public static async Task<object> HttpPost(
        [ExcelArgument(Name = "адрес", Description = "Url адрес, по которому необходимо сделать запрос")]
        string адрес,
        [ExcelArgument(Name = "jsonPath", Description =
            @"Задаётся в формате JSONPath и позволяет сразу перейти к нужному элементу Json.
Подробнее о формате JSONPath по ссылке:
https://danielaparker.github.io/JsonCons.Net/articles/JsonPath/JsonConsJsonPath.html")]
        string jsonPath = "$",
        [ExcelArgument(Name = "заголовки", Description =
            "Диапазон: столбец с названиями и столбцы со значениями заголовка запроса (необязательно)")]
        string[,]? заголовки = null,
        [ExcelArgument(Name = "тело", Description = "Тело запроса в формате json")]
        string? телоJson = null,
        CancellationToken ct = default) {
        // Предотвращает выполнение пока запущен мастер функций
        if (ExcelDnaUtil.IsInFunctionWizard()) return "";

        if (заголовки?.Length == 1)
            заголовки = null;

        var ответ = await HttpКлиент.HttpКлиент.PostАсинх(адрес, заголовки, телоJson, ct).ConfigureAwait(false);

        if (ответ is null)
            return ExcelError.ExcelErrorNA;

        if (string.IsNullOrWhiteSpace(jsonPath) || jsonPath == "$")
            return ответ;

        return JsonКлиент.JsonКлиент.JsonPathНайди(ответ, jsonPath);
    }

    #endregion
}