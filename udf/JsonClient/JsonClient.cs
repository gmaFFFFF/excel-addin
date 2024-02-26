using System.Text.Encodings.Web;
using System.Text.Json;
using JsonCons.JmesPath;
using JsonCons.JsonPath;

namespace gmafffff.excel.udf.JsonКлиент;

public sealed class JsonКлиент {
    private static readonly JsonDocumentOptions JsonДокНастр = new() { AllowTrailingCommas = true };

    private static readonly JsonSerializerOptions НастройкиСохранения = new() {
        WriteIndented = false, AllowTrailingCommas = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };


    /// <summary>
    ///     Находит элемент по JSONPath
    /// </summary>
    /// <param name="jsonТекст">Текст в формате JSON</param>
    /// <param name="запрос">
    ///     Запрос в формате JSONPath.
    ///     Подробнее о формате JSONPath по
    ///     <a href="https://danielaparker.github.io/JsonCons.Net/articles/JsonPath/JsonConsJsonPath.html">ссылке</a>
    /// </param>
    /// <returns>Значение единичного элемента или строку JSON</returns>
    public static object JsonPathНайди(string jsonТекст, string запрос) {
        if (string.IsNullOrWhiteSpace(jsonТекст) || string.IsNullOrWhiteSpace(запрос)) return "";

        var настройкиПоиска = JsonSelectorOptions.Default;
        настройкиПоиска.ExecutionMode = PathExecutionMode.Parallel;

        using var json = JsonDocument.Parse(jsonТекст, JsonДокНастр);

        var найдено = JsonSelector.Select(json.RootElement, запрос, настройкиПоиска);

        return ПреобразоватьJsonВСтроку(найдено);
    }

    public static object JmesPathИзмени(string jsonТекст, string запрос) {
        if (string.IsNullOrWhiteSpace(jsonТекст) || string.IsNullOrWhiteSpace(запрос)) return "";

        var настройкиПоиска = JsonSelectorOptions.Default;
        настройкиПоиска.ExecutionMode = PathExecutionMode.Parallel;


        using var json = JsonDocument.Parse(jsonТекст, JsonДокНастр);

        var найдено = JsonTransformer.Transform(json.RootElement, запрос);

        var элем = найдено is not null
                       ? new[] { найдено.RootElement }
                       : Array.Empty<JsonElement>();

        return ПреобразоватьJsonВСтроку(элем);
    }

    public static object JsonИндекс(string jsonТекст, params string[] индексы) {
        var обработать = (string i) => int.TryParse(i, out _) ? $"[{i}]" : $"['{i}']";
        var запрос     = "$" + string.Concat(индексы.Select(i => обработать(i)));
        return JsonPathНайди(jsonТекст, запрос);
    }

    private static object ПреобразоватьJsonВСтроку(IList<JsonElement> найдено) {
        return найдено.Count switch {
            0 => "",
            1 => ПопробуйТипизироватьЭлементJson(найдено[0]),
            _ => JsonSerializer.Serialize(найдено, НастройкиСохранения)
        };
    }

    private static object ПопробуйТипизироватьЭлементJson(JsonElement элемент) {
        return элемент.ValueKind switch {
                   JsonValueKind.True   => true,
                   JsonValueKind.False  => false,
                   JsonValueKind.Null   => "",
                   JsonValueKind.Number => ПопробуйТипизироватьЭлементJsonКакЧисло(элемент),
                   JsonValueKind.String => элемент.TryGetDateTime(out var дата)
                                               ? дата
                                               : элемент.GetString(),
                   _ => JsonSerializer.Serialize(элемент, НастройкиСохранения)
               }
            ?? "";
    }

    private static object ПопробуйТипизироватьЭлементJsonКакЧисло(JsonElement элемент) {
        if (элемент.TryGetInt64(out var целое)) return целое;
        if (элемент.TryGetDecimal(out var дес)) return дес;
        if (элемент.TryGetDouble(out var вещ)) return вещ;

        return JsonSerializer.Serialize(элемент, НастройкиСохранения);
    }
}