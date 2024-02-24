using System.Text;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace gmafffff.excel.udf.AddIn;

public static class Функции {
    private const string МояКатегория = "Функции от gmaFFFFF";

    #region Форматирование текста

    #region РублиПрописью

    private const string РПИ = nameof(РублиПрописью);
    private const string РПО = "Отображает сумму в рублях прописью";
    private const string РПАСуммаИ = "суммаРублей";
    private const string РПАСуммаО = "сумма, которую необходимо написать прописью";
    private const string РПАФорматИ = "формат";

    private const string РПAФорматО = "ч[n](сумма, n - знаков после запятой), б/д[n][т[з]]" +
                                      "(целая/дробная часть, т - текстом, з - с заглавной, n - ширина), р/к[с]" +
                                      "(валюта базовая/дробная, с - сокращенная). Пример: «ч2 рс (бтз р д2 к)»";

    [ExcelFunction(Name = РПИ, Category = МояКатегория, Description = РПО, IsThreadSafe = true)]
    public static string РублиПрописью(
        [ExcelArgument(Name = РПАСуммаИ, Description = РПАСуммаО)]
        double сумма,
        [ExcelArgument(Name = РПАФорматИ, Description = РПAФорматО)]
        string формат = "") {
        return ДеньгиПрописью.ДеньгиПрописью.РублиПрописью(сумма, формат);
    }

    #endregion

    #region ОкруглГаус

    private const string ОГИ = nameof(ОкруглГаус);
    private const string ОГО = "Округление по Гауссу до ближайшего четного знака";
    private const string ОГАЧислоИ = "число";
    private const string ОГАЧислоО = "округляемое число";
    private const string ОГАЗнаковИ = "знаков";

    private const string ОГАЗнаковО = "число знаков, до которых происходит округление. Если < 0, то перед запятой. " +
                                      "Максимум 15 знаков, по умолчанию — 0";

    [ExcelFunction(Name = ОГИ, Category = МояКатегория, Description = ОГО, IsThreadSafe = true)]
    public static double ОкруглГаус(
        [ExcelArgument(Name = ОГАЧислоИ, Description = ОГАЧислоО)]
        double число,
        [ExcelArgument(Name = ОГАЗнаковИ, Description = ОГАЗнаковО)]
        short знаков) {
        return (знаков, Math.Pow(10, -знаков)) switch {
            (> 15, _) => Math.Round(число, 15, MidpointRounding.ToEven),
            (>= 0, _) => Math.Round(число, знаков, MidpointRounding.ToEven),
            (< 0, var степень) => Math.Round(число / степень, 0, MidpointRounding.ToEven) * степень
        };
    }

    #endregion

    #endregion

    #region Управляющие функции

    #region Coalesce

    private const string CoalИ = nameof(Coalesce);
    private const string CoalО = "Возвращает первый из аргументов, не являющихся ошибкой или пустым";
    private const string CoalАИ = "аргумент";
    private const string CoalАО = "проверяемый аргумент";

    [ExcelFunction(Name = CoalИ, Category = МояКатегория, Description = CoalО, IsThreadSafe = true)]
    public static object Coalesce(
        [ExcelArgument(Name = CoalАИ, Description = CoalАО)]
        params object?[]? список) {
        if (список is null)
            return ExcelError.ExcelErrorNull;

        object рез = ExcelError.ExcelErrorNull;

        foreach (var элем in список) {
            if (элем is Array ar)
                рез = (from object з in ar
                        where ЗначимЛи(з)
                        select з)
                    .FirstOrDefault(ExcelError.ExcelErrorNull);
            else
                рез = ЗначимЛи(элем) ? элем : рез;

            if (рез is not ExcelError)
                return рез!;
        }

        return рез;

        bool ЗначимЛи(object? o) => o is not null && o is not ExcelError && o is not ExcelMissing &&
                                    o is not ExcelEmpty && o is string s && !string.IsNullOrEmpty(s);
    }

    #endregion

    #region Скрыть строку/столбец

    private const string ВидСтрИ = nameof(ОтобрСтр);
    private const string ВидСтрО = "Скрывает/отображает строку в зависимости от значения параметра";
    private const string ВидСтрПереклИ = "видимаЛи";
    private const string ВидСтрПереклО = "Истина() — строка видна, Ложь() — строка скрыта";
    private const string ВидСтрСсылкаИ = "ссылка";
    private const string ВидСтрСсылкаО = "укажи строку";
    private const string ВидСтрВысотаИ = "высота";

    private const string ВидСтрВысотаО = "необязательная высота отображенной строки, " +
                                         "соответствует высоте шрифта по умолчанию";

    [ExcelFunction(Name = ВидСтрИ, Category = МояКатегория, Description = ВидСтрО, IsMacroType = true)]
    public static object ОтобрСтр(
        [ExcelArgument(Name = ВидСтрПереклИ, Description = ВидСтрПереклО)]
        bool видимаЛи,
        [ExcelArgument(Name = ВидСтрСсылкаИ, Description = ВидСтрСсылкаО, AllowReference = true)]
        object парам,
        [ExcelArgument(Name = ВидСтрВысотаИ, Description = ВидСтрВысотаО)]
        double? высота = null
    ) {
        if (парам is not ExcelReference ссылка)
            return ExcelError.ExcelErrorRef;

        // UDF должны выполняться без побочных эффектов
        // Данная функция нарушает данное правило, но делает это аккуратно (если это вообще возможно) —
        // побочный эффект будет поставлен в очередь основного потока Excel, когда он будет готов
        foreach (var стр in ExcelСтрока.Преобразовать(ссылка)) {
            var команда = new ExcelКомандаВидимости(стр, видимаЛи, высота);
            ExcelМенеджерФоновыхКоманд.ЗапланироватьКоманду(команда);
        }

        return видимаЛи;
    }


    private const string ВидСтлбИ = nameof(ОтобрСтлб);
    private const string ВидСтлбО = "Скрывает/отображает столбец в зависимости от значения параметра";
    private const string ВидСтлбПереклИ = "видимЛи";
    private const string ВидСтлбПереклО = "Истина() — столбец виден, Ложь() — столбец скрыт";
    private const string ВидСтлбСсылкаИ = "ссылка";
    private const string ВидСтлбСсылкаО = "укажи столбец";
    private const string ВидСтлбШиринаИ = "ширина";

    private const string ВидСтлбШиринаО = "необязательная ширина отображенного столбца, " +
                                          "соответствует ширине символа шрифта по умолчанию";

    [ExcelFunction(Name = ВидСтлбИ, Category = МояКатегория, Description = ВидСтлбО, IsMacroType = true)]
    public static object ОтобрСтлб(
        [ExcelArgument(Name = ВидСтлбПереклИ, Description = ВидСтлбПереклО)]
        bool видимЛи,
        [ExcelArgument(Name = ВидСтлбСсылкаИ, Description = ВидСтлбСсылкаО, AllowReference = true)]
        object парам,
        [ExcelArgument(Name = ВидСтлбШиринаИ, Description = ВидСтлбШиринаО, AllowReference = true)]
        double? ширина = null
    ) {
        // Предотвращает выполнение пока запущен мастер функций
        if (ExcelDnaUtil.IsInFunctionWizard()) return "";

        if (парам is not ExcelReference ссылка)
            return ExcelError.ExcelErrorRef;

        // UDF должны выполняться без побочных эффектов
        // Данная функция нарушает данное правило, но делает это аккуратно (если это вообще возможно) —
        // побочный эффект будет поставлен в очередь основного потока Excel, когда он будет готов
        foreach (var стр in ExcelСтолбец.Преобразовать(ссылка)) {
            var команда = new ExcelКомандаВидимости(стр, видимЛи, ширина);
            ExcelМенеджерФоновыхКоманд.ЗапланироватьКоманду(команда);
        }

        return видимЛи;
    }

    #endregion

    #endregion

    #region Http

    #region Get/Post

    private const string HGАктивИ = nameof(HttpGet_active);
    private const string HPАктивИ = nameof(HttpPost_active);

    private const string HGPОобщее = "запрос возвращает json с полями: " +
                                     $"{nameof(HttpКлиент.HttpКлиент.ОтветHttp.Статус)}; " +
                                     $"{nameof(HttpКлиент.HttpКлиент.ОтветHttp.ДатаЗапроса)}; " +
                                     $"{nameof(HttpКлиент.HttpКлиент.ОтветHttp.Содержимое)}; " +
                                     $"{nameof(HttpКлиент.HttpКлиент.ОтветHttp.Заголовки)} (ответа); " +
                                     $"{nameof(HttpКлиент.HttpКлиент.ОтветHttp.Заголовки2)} (содержимого); " +
                                     $"{nameof(HttpКлиент.HttpКлиент.ОтветHttp.Куки)}.\n";

    private const string HGPАктив_Предупреждение = "Внимание: " +
                                                   "1)выполняется ОЧЕНЬ часто — при пересчете листа," +
                                                   "2)кол-во символов в ячейке ограничено";

    private const string HGАктивО = "Get " + HGPОобщее + HGPАктив_Предупреждение;
    private const string HPАктивО = "Post " + HGPОобщее + HGPАктив_Предупреждение;

    private const string HGPААдресИ = "адрес";
    private const string HGPААдресО = "Url адрес, по которому необходимо сделать запрос";
    private const string HGPАJsonPathИ = "jsonPath";

    private const string HGPАJsonPathО = "Необязательный JSONPath позволяет выбрать нужный элемент из ответа.\n" +
                                         "Подробнее о формате JSONPath по ссылке:\n" +
                                         JsonPathHelpUrl;

    private const string HGPАЗаголовкиИ = "заголовки";
    private const string HGPАЗаголовкиJИ = "заголовкиJSON";

    private const string HGPАЗаголовкиО = "Необязательный диапазон с заголовками запроса:\n" +
                                          "один столбец с заголовком\n" +
                                          "один или несколько столбцов со значениями заголовка";

    private const string HGPАЗаголовкиJО = "Необязательные заголовки в формате массива объектов JSON:\n" +
                                           "[{\"Заголовок1\":[\"знач1\",\"знач2\"]},{\"Заголовок2\":\"знач3\"}]";

    private const string HPАТелоИ = "тело";
    private const string HPАТелоО = "Тело запроса в формате json (необязательно)";

    [ExcelAsyncFunction(Name = HGАктивИ, Category = МояКатегория, Description = HGАктивО)]
    public static async Task<object> HttpGet_active(
        [ExcelArgument(Name = HGPААдресИ, Description = HGPААдресО)]
        string адрес,
        [ExcelArgument(Name = HGPАJsonPathИ, Description = HGPАJsonPathО)]
        string? jsonPath = null,
        [ExcelArgument(Name = HGPАЗаголовкиИ, Description = HGPАЗаголовкиО)]
        string[,]? заголовки = null,
        [ExcelArgument(Name = HGPАЗаголовкиJИ, Description = HGPАЗаголовкиJО)]
        string? заголовкиJson = null,
        CancellationToken ct = default) {
        // Предотвращает выполнение пока запущен мастер функций
        if (ExcelDnaUtil.IsInFunctionWizard()) return "";

        if (заголовки?.Length == 1)
            заголовки = null;

        var ответ = await HttpКлиент.HttpКлиент.GetАсинх(адрес, заголовки, заголовкиJson, ct).ConfigureAwait(false);

        if (ответ is null)
            return ExcelError.ExcelErrorNA;

        if (string.IsNullOrWhiteSpace(jsonPath) || jsonPath == "$")
            return ответ;

        return JsonКлиент.JsonКлиент.JsonPathНайди(ответ, jsonPath);
    }

    [ExcelAsyncFunction(Name = HPАктивИ, Category = МояКатегория, Description = HPАктивО)]
    public static async Task<object> HttpPost_active(
        [ExcelArgument(Name = HGPААдресИ, Description = HGPААдресО)]
        string адрес,
        [ExcelArgument(Name = HGPАJsonPathИ, Description = HGPАJsonPathО)]
        string? jsonPath = null,
        [ExcelArgument(Name = HGPАЗаголовкиИ, Description = HGPАЗаголовкиО)]
        string[,]? заголовки = null,
        [ExcelArgument(Name = HGPАЗаголовкиJИ, Description = HGPАЗаголовкиJО)]
        string? заголовкиJson = null,
        [ExcelArgument(Name = HPАТелоИ, Description = HPАТелоО)]
        string? телоJson = null,
        CancellationToken ct = default) {
        // Предотвращает выполнение пока запущен мастер функций
        if (ExcelDnaUtil.IsInFunctionWizard()) return "";

        if (заголовки?.Length == 1)
            заголовки = null;

        var ответ = await HttpКлиент.HttpКлиент.PostАсинх(адрес, заголовки, заголовкиJson, телоJson, ct)
            .ConfigureAwait(false);

        if (ответ is null)
            return ExcelError.ExcelErrorNA;

        if (string.IsNullOrWhiteSpace(jsonPath) || jsonPath == "$")
            return ответ;

        return JsonКлиент.JsonКлиент.JsonPathНайди(ответ, jsonPath);
    }

    private const string HGifИ = nameof(HttpGet_if);
    private const string HPifИ = nameof(HttpPost_if);

    private const string HGPif_Предупреждение = "Внимание: " +
                                                "1)Результат кешируется в самой ячейке — нельзя вкладывать в другую функцию," +
                                                "2)кол-во символов в ячейке ограничено";

    private const string HGifО = "Get " + HGPОобщее + HGPif_Предупреждение;
    private const string HPifО = "Post " + HGPОобщее + HGPif_Предупреждение;
    private const string HGPifПересчетИ = "повторитьЛи";
    private const string HGPifПересчетО = "нужно ли повторно выполнить запрос или использовать кеш";
    private const string HGPifJmesPathИ = "JMESPath";

    private const string HGPifJmesPathО = "Необязательный JMESPath позволяет выбрать нужный элемент из ответа.\n" +
                                          "Подробнее о формате JMESPath по ссылке:\n" +
                                          JmesPathHelpUrl;

    [ExcelAsyncFunction(Name = HGifИ, Category = МояКатегория, Description = HGifО, IsMacroType = true)]
    public static async Task<string> HttpGet_if(
        [ExcelArgument(Name = HGPifПересчетИ, Description = HGPifПересчетО)]
        bool повторитьЛи,
        [ExcelArgument(Name = HGPААдресИ, Description = HGPААдресО)]
        string адрес,
        [ExcelArgument(Name = HGPifJmesPathИ, Description = HGPifJmesPathО)]
        string? jmesPath = null,
        [ExcelArgument(Name = HGPАЗаголовкиИ, Description = HGPАЗаголовкиО)]
        string[,]? заголовки = null,
        [ExcelArgument(Name = HGPАЗаголовкиJИ, Description = HGPАЗаголовкиJО)]
        string? заголовкиJson = null,
        CancellationToken ct = default) {
        // Предотвращает выполнение пока запущен мастер функций
        if (ExcelDnaUtil.IsInFunctionWizard()) return "";

        var запросиИВерниФунк = async () => {
            var ответ = await HttpGet_active(адрес, null, заголовки, заголовкиJson, ct).ConfigureAwait(false) as string;
            if (ответ is not { } str) return ExcelError.ExcelErrorNA.ToString();

            return string.IsNullOrWhiteSpace(jmesPath)
                ? str
                : JsonКлиент.JsonКлиент.JmesPathИзмени(str, jmesPath).ToString() ?? ExcelError.ExcelErrorNA.ToString();
        };

        if (повторитьЛи || XlCall.Excel(XlCall.xlfCaller) is not ExcelReference вызванИз)
            return await запросиИВерниФунк();

        var value = вызванИз.GetValue();
        if (value is (not ExcelError or ExcelMissing) and string old && !string.IsNullOrWhiteSpace(old))
            return await Task.FromResult(old);

        return await запросиИВерниФунк();
    }

    [ExcelAsyncFunction(Name = HPifИ, Category = МояКатегория, Description = HPifО, IsMacroType = true)]
    public static async Task<string?> HttpPost_if(
        [ExcelArgument(Name = HGPifПересчетИ, Description = HGPifПересчетО)]
        bool повторитьЛи,
        [ExcelArgument(Name = HGPААдресИ, Description = HGPААдресО)]
        string адрес,
        [ExcelArgument(Name = HGPifJmesPathИ, Description = HGPifJmesPathО)]
        string? jmesPath = null,
        [ExcelArgument(Name = HGPАЗаголовкиИ, Description = HGPАЗаголовкиО)]
        string[,]? заголовки = null,
        [ExcelArgument(Name = HGPАЗаголовкиJИ, Description = HGPАЗаголовкиJО)]
        string? заголовкиJson = null,
        [ExcelArgument(Name = HPАТелоИ, Description = HPАТелоО)]
        string? телоJson = null,
        CancellationToken ct = default) {
        // Предотвращает выполнение пока запущен мастер функций
        if (ExcelDnaUtil.IsInFunctionWizard()) return "";

        var запросиИВерниФунк = async () => {
            var ответ =
                await HttpPost_active(адрес, null, заголовки, заголовкиJson, телоJson, ct)
                    .ConfigureAwait(false) as string;
            if (ответ is not { } str) return ExcelError.ExcelErrorNA.ToString();

            return string.IsNullOrWhiteSpace(jmesPath)
                ? str
                : JsonКлиент.JsonКлиент.JmesPathИзмени(str, jmesPath).ToString() ?? ExcelError.ExcelErrorNA.ToString();
        };

        if (повторитьЛи || XlCall.Excel(XlCall.xlfCaller) is not ExcelReference вызванИз)
            return await запросиИВерниФунк();

        var value = вызванИз.GetValue();
        if (value is (not ExcelError or ExcelMissing) and string old && !string.IsNullOrWhiteSpace(old))
            return await Task.FromResult(old);

        return await запросиИВерниФунк();
    }

    #endregion

    #region Json // JSONPath, JMESPath, JsonИндекс

    private const string JsonPathHelpUrl =
        "https://danielaparker.github.io/JsonCons.Net/articles/JsonPath/JsonConsJsonPath.html";

    private const string JmesPathHelpUrl = "https://jmespath.org/specification.html";

    private const string JИИ = nameof(JsonИндекс);
    private const string JPИ = nameof(JsonPath);
    private const string JmPИ = nameof(JmesPath);

    private const string JИО = "Извлекает элементы json по индексу";

    private const string JPО = "Извлекает элементы json с помощью синтаксиса JSONPath. " +
                               "Не умеет проецировать данные (например, фильтрация с последующим индексом массива). " +
                               "При необходимости проецировать данные используйте функцию JmesPath.\n" +
                               "Примеры запросов в справке";

    private const string JmPО = "Извлекает элементы json с помощью синтаксиса JMESPath.\n" +
                                "Примеры запросов в справке";

    private const string JPJMАJsonТекстИ = "jsonТекст";
    private const string JPJMАJsonТекстО = "Json документ";

    private const string JИАИндексИ = "индекс";
    private const string JИАИндексО = "индексы для доступа к JSON";
    private const string JPАJsonPathИ = "jsonPath";
    private const string JPАJsonPathО = "JSONPath, подробнее о формате в справке к функции";
    private const string JmPАJsonPathИ = "jmesPath";
    private const string JmPАJsonPathО = "JMESPath, подробнее о формате в справке к функции";

    [ExcelFunction(Name = JИИ, Category = МояКатегория, Description = JИО, IsThreadSafe = true)]
    public static object JsonИндекс(
        [ExcelArgument(Name = JPJMАJsonТекстИ, Description = JPJMАJsonТекстО)]
        string jsonТекст,
        [ExcelArgument(Name = JИАИндексИ, Description = JИАИндексО)]
        params string[] индексы) 
            => JsonКлиент.JsonКлиент.JsonИндекс(jsonТекст, индексы);

    [ExcelFunction(Name = JPИ, Category = МояКатегория, Description = JPО, HelpTopic = JsonPathHelpUrl,
        IsThreadSafe = true)]
    public static object JsonPath(
        [ExcelArgument(Name = JPJMАJsonТекстИ, Description = JPJMАJsonТекстО)]
        string jsonТекст,
        [ExcelArgument(Name = JPАJsonPathИ, Description = JPАJsonPathО)]
        string jsonPath) {
        return JsonКлиент.JsonКлиент.JsonPathНайди(jsonТекст, jsonPath);
    }

    [ExcelFunction(Name = JmPИ, Category = МояКатегория, Description = JmPО, HelpTopic = JmesPathHelpUrl,
        IsThreadSafe = true)]
    public static object JmesPath(
        [ExcelArgument(Name = JPJMАJsonТекстИ, Description = JPJMАJsonТекстО)]
        string jsonТекст,
        [ExcelArgument(Name = JmPАJsonPathИ, Description = JmPАJsonPathО)]
        string jmesPath) {
        return JsonКлиент.JsonКлиент.JmesPathИзмени(jsonТекст, jmesPath);
    }

    #endregion

    #region Base64

    // Base64Кодировать и Base64Декодировать
    private const string B64КИ = nameof(Base64Кодировать);
    private const string B64ДИ = nameof(Base64Декодировать);

    private const string B64КО = "Кодирует текст в base64 код";
    private const string B64ДО = "Декодирует текст из формата base64";

    private const string B64КАТекстИ = "текст";
    private const string B64КАТекстО = "текст для кодирования";

    private const string B64ДАbase64ТекстИ = "base64Текст";
    private const string B64ДАbase64ТекстО = "текст закодированный в формате base64";

    [ExcelFunction(Name = B64КИ, Category = МояКатегория, Description = B64КО, IsThreadSafe = true)]
    public static string Base64Кодировать(
        [ExcelArgument(Name = B64КАТекстИ, Description = B64КАТекстО)]
        string текст) {
        var байты = Encoding.UTF8.GetBytes(текст);
        return Convert.ToBase64String(байты);
    }

    [ExcelFunction(Name = B64ДИ, Category = МояКатегория, Description = B64ДО, IsThreadSafe = true)]
    public static string Base64Декодировать(
        [ExcelArgument(Name = B64ДАbase64ТекстИ, Description = B64ДАbase64ТекстО)]
        string base64Текст) {
        var байты = Convert.FromBase64String(base64Текст);
        return Encoding.UTF8.GetString(байты);
    }

    #endregion

    #endregion
}