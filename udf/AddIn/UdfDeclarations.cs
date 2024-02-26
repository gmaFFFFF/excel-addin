using System.Collections;
using System.Globalization;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Registration;
using gmafffff.excel.udf.Excel.Команды;
using gmafffff.excel.udf.Excel.Сетка;

namespace gmafffff.excel.udf.AddIn;

public static class Функции{
    private const string МояКатегория = "Функции от gmaFFFFF";
    private static readonly ИОчередьКоманд ОчередьКоманд = new ОчередьКомандRx();

    static Функции(){
        ОчередьКоманд.ДобавитьКомпоновщикКоманд(typeof(ИзмениВидимостьРядаКоманда),
            ИзмениВидимостьРядаКоманда.Упаковать);
    }

    #region Служебные

    private static bool ЗначимЛиАргументUdf(object? o) {
        return !(o is null || o is ExcelError || o is ExcelMissing ||
                 o is ExcelEmpty || (o is string s && string.IsNullOrEmpty(s)));
    }


    private static IEnumerable<object?> FlattenUdfArgument(object?[]? список){
        if (список is null) yield break;
        foreach (var элем in список)
            if (элем is Array ar) {
                var массив = from object з in ar
                    select з;
                foreach (var элемМассива in массив)
                    yield return элемМассива;
            }
            else {
                yield return элем;
            }
    }

    #endregion

    #region Форматирование текста

    #region «Интерполяция строк»

    private const string НаборСтрИ = nameof(НаборСтроки);

    private const string НаборСтрО =
        "Замена заполнителей в строке ({0}, {1}) переданными в качестве аргументов функции значениями";

    private const string НаборСтрHelp =
        "https://learn.microsoft.com/ru-ru/dotnet/standard/base-types/composite-formatting";

    private const string НаборСтрСтрИ = "текст";
    private const string НаборСтрСтрО = "Текст с заполнителями {0}, {1}";
    private const string НаборСтрЗначИ = "значения";
    private const string НаборСтрЗначО = "Значения заполнителей";

    [ExcelFunction(Name = НаборСтрИ, Category = МояКатегория, Description = НаборСтрО, IsThreadSafe = true,
        HelpTopic = НаборСтрHelp)]
    public static object НаборСтроки(
        [ExcelArgument(Name = НаборСтрСтрИ, Description = НаборСтрСтрО)]
        string text,
        [ExcelArgument(Name = НаборСтрЗначИ, Description = НаборСтрЗначО)]
        params object[] значения) {
        var ошибки = from з in значения
            where з is ExcelError || з is ExcelMissing
            select з;
        if (ошибки.Any()) return ExcelError.ExcelErrorNA;
        
        return string.Format(text, значения);
    }

    #endregion

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
        string формат = ""){
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
        short знаков){
        return (знаков, Math.Pow(10, -знаков)) switch
        {
            (> 15, _) => Math.Round(число, 15, MidpointRounding.ToEven),
            (>= 0, _) => Math.Round(число, знаков, MidpointRounding.ToEven),
            (< 0, var степень) => Math.Round(число / степень, 0, MidpointRounding.ToEven) * степень
        };
    }

    #endregion

    #region Сократить фамилию

    private const string СФиоИ = nameof(СократитьФио);
    private const string СФиоО = "Сокращает полные Фамилия Имя Отчество до ФИО";
    private const string СФиоАФиоИ = "ФИО";
    private const string СФиоАФиоО = "сокращаемая Фамилия Имя Отчество";
    private const string СФиоАСлеваИ = "слева";
    private const string СФиоАСлеваО = "инициалы слева?";

    [ExcelFunction(Name = СФиоИ, Category = МояКатегория, Description = СФиоО, IsThreadSafe = true)]
    public static string СократитьФио(
        [ExcelArgument(Name = СФиоАФиоИ, Description = СФиоАФиоО)]
        string фио,
        [ExcelArgument(Name = СФиоАСлеваИ, Description = СФиоАСлеваО)]
        bool слева = false){
        return ФИО.ФИО.СократитьФио(фио, слева);
    }

    #endregion

    #region Реверс строки

    private const string РеверсИ = nameof(Реверс);
    private const string РеверсО = "Возвращает символы текста в обратном порядке";
    private const string РеверсСтрИ = "текст";
    private const string РеверсСтрО = "Текст для реверса";

    [ExcelFunction(Name = РеверсИ, Category = МояКатегория, Description = РеверсО, IsThreadSafe = true)]
    public static string Реверс(
        [ExcelArgument(Name = РеверсСтрИ, Description = РеверсСтрО)]
        string text){
        return string.Join("", GetTextElements(text)
            .Reverse()
            .ToArray());

        IEnumerable<string> GetTextElements(string? text){
            // Источник: https://stackoverflow.com/a/15111719
            var enumerator = StringInfo.GetTextElementEnumerator(text ?? string.Empty);
            while (enumerator.MoveNext()) yield return (enumerator.Current as string)!;
        }
    }

    #endregion

    #region Первая прописаня

    private const string Прописная1И = nameof(ПрописнаяПервая);
    private const string Прописная1О = "Делает первую букву в строке прописной";
    private const string Прописная1СтрИ = "текст";
    private const string Прописная1СтрО = "Текст";

    [ExcelFunction(Name = Прописная1И, Category = МояКатегория, Description = Прописная1О, IsThreadSafe = true)]
    public static string ПрописнаяПервая(
        [ExcelArgument(Name = Прописная1СтрИ, Description = Прописная1СтрО)]
        string text){
        return $"{text[0].ToString().ToUpper()}{text[1..]}";
    }

    #endregion

    #endregion

    #region Информация

    #region Информация о пользователе

    private const string ТПИ = nameof(ТекущийПользователь);
    private const string ТПО = "Доступная информация о текущем пользователе из ActiveDirectory";

    [ExcelFunction(Name = ТПИ, Category = МояКатегория, Description = ТПО, IsThreadSafe = true)]
    public static string ТекущийПользователь(){
        return new User.User().Json();
    }

    #endregion

    #endregion

    #region Файловые функции

    #region Проверка существования файла

    private const string ФайлСущЛиИ = nameof(ФайлСуществуетЛи);
    private const string ФайлСущЛиО = "Проверяет существование файла";
    private const string ФайлСущЛиПутьИ = "путь";
    private const string ФайлСущЛиПутьО = "путь к файлу";

    [ExcelFunction(Name = ФайлСущЛиИ, Category = МояКатегория, Description = ФайлСущЛиО, IsThreadSafe = true)]
    public static bool ФайлСуществуетЛи(
        [ExcelArgument(Name = ФайлСущЛиПутьИ, Description = ФайлСущЛиПутьО)]
        string path){
        return File.Exists(path);
    }

    #endregion

    #endregion

    #region Управляющие функции

    #region Coalesce

    private const string CoalИ = nameof(Coalesce);

    private const string CoalО = "Возвращает первый из аргументов, не являющихся ошибкой или пустым." +
                                 "Если такого элемента нет, то возвращается пустая строка";

    private const string CoalАИ = "аргумент";
    private const string CoalАО = "проверяемый аргумент";

    [ExcelFunction(Name = CoalИ, Category = МояКатегория, Description = CoalО, IsThreadSafe = true)]
    public static object Coalesce(
        [ExcelArgument(Name = CoalАИ, Description = CoalАО)]
        params object?[]? список){
        if (список is null || !((object[])список).Any())
            return ExcelError.ExcelErrorNull;
        var аргументыUdf = FlattenUdfArgument(список);

        return (from арг in аргументыUdf
                where ЗначимЛиАргументUdf(арг)
                select арг)
            .FirstOrDefault("");
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
        double? высота = null){
        if (парам is not ExcelReference ссылка)
            return ExcelError.ExcelErrorRef;

        // UDF должны выполняться без побочных эффектов.
        // Данная функция нарушает данное правило, но делает это аккуратно (если это вообще возможно) —
        // побочный эффект будет поставлен в очередь основного потока Excel, когда он будет свободен
        foreach (var стр in ExcelСтрока.Преобразовать(ссылка)) {
            var команда = new ИзмениВидимостьРядаКоманда(стр, видимаЛи, высота);
            ОчередьКоманд.ДобавитьКоманду(команда);
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
        double? ширина = null){
        // Предотвращает выполнение пока запущен мастер функций
        if (ExcelDnaUtil.IsInFunctionWizard()) return "";

        if (парам is not ExcelReference ссылка)
            return ExcelError.ExcelErrorRef;

        // UDF должны выполняться без побочных эффектов.
        // Данная функция нарушает данное правило, но делает это аккуратно (если это вообще возможно) —
        // побочный эффект будет поставлен в очередь основного потока Excel, когда он будет готов
        foreach (var стр in ExcelСтолбец.Преобразовать(ссылка)) {
            var команда = new ИзмениВидимостьРядаКоманда(стр, видимЛи, ширина);
            ОчередьКоманд.ДобавитьКоманду(команда);
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
        CancellationToken ct = default){
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
        CancellationToken ct = default){
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
        CancellationToken ct = default){
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
        CancellationToken ct = default){
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
        params string[] индексы){
        return JsonКлиент.JsonКлиент.JsonИндекс(jsonТекст, индексы);
    }

    [ExcelFunction(Name = JPИ, Category = МояКатегория, Description = JPО, HelpTopic = JsonPathHelpUrl,
        IsThreadSafe = true)]
    public static object JsonPath(
        [ExcelArgument(Name = JPJMАJsonТекстИ, Description = JPJMАJsonТекстО)]
        string jsonТекст,
        [ExcelArgument(Name = JPАJsonPathИ, Description = JPАJsonPathО)]
        string jsonPath){
        return JsonКлиент.JsonКлиент.JsonPathНайди(jsonТекст, jsonPath);
    }

    [ExcelFunction(Name = JmPИ, Category = МояКатегория, Description = JmPО, HelpTopic = JmesPathHelpUrl,
        IsThreadSafe = true)]
    public static object JmesPath(
        [ExcelArgument(Name = JPJMАJsonТекстИ, Description = JPJMАJsonТекстО)]
        string jsonТекст,
        [ExcelArgument(Name = JmPАJsonPathИ, Description = JmPАJsonPathО)]
        string jmesPath){
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
        string текст){
        var байты = Encoding.UTF8.GetBytes(текст);
        return Convert.ToBase64String(байты);
    }

    [ExcelFunction(Name = B64ДИ, Category = МояКатегория, Description = B64ДО, IsThreadSafe = true)]
    public static string Base64Декодировать(
        [ExcelArgument(Name = B64ДАbase64ТекстИ, Description = B64ДАbase64ТекстО)]
        string base64Текст){
        var байты = Convert.FromBase64String(base64Текст);
        return Encoding.UTF8.GetString(байты);
    }

    #endregion

    #endregion

    #region Функции массива

    private static bool ЧисловойЛи(this object o){
        var numericTypes = new HashSet<Type>
        {
            //встроенные:
            typeof(byte),
            typeof(sbyte),
            typeof(ushort),
            typeof(uint),
            typeof(ulong),
            typeof(short),
            typeof(int),
            typeof(long),
            typeof(decimal),
            typeof(double),
            typeof(float)
        };
        return numericTypes.Contains(o.GetType());
    }

    private class ЦифрыПередТекстомСравниватель : IComparer, IComparer<object?>{
        // Call CaseInsensitiveComparer.Compare with the parameters reversed.
        int IComparer.Compare(object? x, object? y){
            return (x, y) switch
            {
                (null, not null) => -1,
                (null, null) => 0,
                (not null, null) => 1,
                (bool b1, bool b2) => b1.CompareTo(b2),
                (bool, _) => -1,
                (_, bool) => 1,
                _ => (x.ЧисловойЛи(), y.ЧисловойЛи()) switch
                {
                    (true, true) => Convert.ToDouble(x).CompareTo(Convert.ToDouble(y)),
                    (true, false) => -1,
                    (false, true) => 1,
                    _ => x.ToString()!.CompareTo(y.ToString())
                }
            };
        }

        int IComparer<object>.Compare(object? x, object? y){
            return ((IComparer)this).Compare(x, y);
        }
    }

    #region Соединить списки

    private const string СоедСписИ = nameof(СоединитьСписки);
    private const string СоедСписО = "Соединяет списки в один";
    private const string СоедСписСИ = "списки";
    private const string СоедСписСО = "Объединяемые списки";

    [ExcelFunction(Name = СоедСписИ, Category = МояКатегория, Description = СоедСписО, IsThreadSafe = true)]
    public static object?[] СоединитьСписки(
        [ExcelArgument(Name = СоедСписСИ, Description = СоедСписСО)]
        params object?[]? списки){
        return FlattenUdfArgument(списки).ToArray();
    }

    #endregion

    #region Сортировать список

    private const string СортирСписИ = nameof(Сортировать);
    private const string СортирСписО = "Сортирует элементы в списке";
    private const string СортирСписСИ = "список";
    private const string СортирСписСО = "Сортируемый список";

    private const string СортирСписПИ = "поУбыванию";

    private const string СортирСписПО = "По умолчанию сортировка идет по возростанию (Ложь)";

    [ExcelFunction(Name = СортирСписИ, Category = МояКатегория, Description = СортирСписО, IsThreadSafe = true)]
    public static object?[] Сортировать(
        [ExcelArgument(Name = СортирСписСИ, Description = СортирСписСО)]
        object?[]? списки,
        [ExcelArgument(Name = СортирСписПИ, Description = СортирСписПО)]
        bool поУбыванию = false){
        return поУбыванию
            ? FlattenUdfArgument(списки).OrderByDescending(e => e, new ЦифрыПередТекстомСравниватель()).ToArray()
            : FlattenUdfArgument(списки).OrderBy(e => e, new ЦифрыПередТекстомСравниватель()).ToArray();
    }

    #endregion

    #region Убрать повторы в списке

    private const string УникСписИ = nameof(УбратьПовторы);
    private const string УникСписО = "Оставляет только уникальные элементы в списке";
    private const string УникСписСИ = "список";
    private const string УникСписСО = "Список, в котором содержаться повторы";

    [ExcelFunction(Name = УникСписИ, Category = МояКатегория, Description = УникСписО, IsThreadSafe = true)]
    public static object?[] УбратьПовторы(
        [ExcelArgument(Name = УникСписСИ, Description = УникСписСО)]
        params object?[]? списки){
        return FlattenUdfArgument(списки).Distinct().ToArray();
    }

    #endregion

    #region Оставить только значимые элементы в списке

    private const string ОставитьЗначИ = nameof(ОставитьЗначимые);
    private const string ОставитьЗначО = "Убирает из списка не значимые элементы: ошибки, пустые строки";
    private const string ОставитьЗначСИ = "список";
    private const string ОставитьЗначСО = "Список, в котором содержаться не значимые элементы";

    [ExcelFunction(Name = ОставитьЗначИ, Category = МояКатегория, Description = ОставитьЗначО, IsThreadSafe = true)]
    public static object?[] ОставитьЗначимые(
        [ExcelArgument(Name = ОставитьЗначСИ, Description = ОставитьЗначСО)]
        params object?[]? списки){
        return FlattenUdfArgument(списки).Where(ЗначимЛиАргументUdf).ToArray();
    }

    #endregion

    #region Число значимых элементов в списке

    private const string ЧислоЗначИ = nameof(ЧислоЗначимых);
    private const string ЧислоЗначО = "Возвращает количество значимых элементов (не ошибки, и не пустые строки)";
    private const string ЧислоЗначСИ = "список";
    private const string ЧислоЗначСО = "Список, в котором содержаться не значимые элементы";

    [ExcelFunction(Name = ЧислоЗначИ, Category = МояКатегория, Description = ЧислоЗначО, IsThreadSafe = true)]
    public static int ЧислоЗначимых(
        [ExcelArgument(Name = ЧислоЗначСИ, Description = ЧислоЗначСО)]
        params object?[]? списки){
        return FlattenUdfArgument(списки).Where(ЗначимЛиАргументUdf).Count();
    }

    #endregion

    #endregion
}