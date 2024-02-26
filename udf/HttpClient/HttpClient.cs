using System.Net;
using System.Net.Http.Headers;
using System.Net.Mime;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace gmafffff.excel.udf.HttpКлиент;

// Представляет любой объект в формате JSON
using ОбъектJson = Dictionary<string, JsonElement>;

public sealed class HttpКлиент {
    private static readonly HttpClient httpКлиент;

    static HttpКлиент() {
        var handler = new SocketsHttpHandler();
        handler.PooledConnectionLifetime = TimeSpan.FromSeconds(60); //Как долго соединение может удерживаться открытым
        handler.PooledConnectionIdleTimeout =
            TimeSpan.FromSeconds(100); //Время бездействия, после которого соединение будет закрыто

        httpКлиент = new HttpClient();
    }

    public static async Task<string?> GetАсинх(string адрес, string[,]? заголовкиСтр = null,
                                               string? заголовкиJson = null,
                                               CancellationToken ct = default) {
        return await ПослатьАсинх(HttpMethod.Get, адрес, заголовкиСтр ?? new string[0, 0], заголовкиJson, null, ct)
                  .ConfigureAwait(false);
    }

    public static async Task<string?> PostАсинх(string адрес, string[,]? заголовкиСтр,
                                                string? заголовкиJson = null, string? телоJson = null,
                                                CancellationToken ct = default) {
        return await ПослатьАсинх(HttpMethod.Post, адрес, заголовкиСтр ?? new string[0, 0], заголовкиJson, телоJson, ct)
                  .ConfigureAwait(false);
    }

    public static async Task<string?> ПослатьАсинх(HttpMethod метод, string адрес, string[,] заголовкиСтр,
                                                   string? заголовкиJson, string? телоJson,
                                                   CancellationToken ct = default) {
        if (string.IsNullOrWhiteSpace(адрес)) return null;

        var jsonSerializerOptions = new JsonSerializerOptions {
            WriteIndented = false, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        телоJson = string.IsNullOrWhiteSpace(телоJson) ? null : телоJson.Trim();
        адрес    = адрес.Trim();

        using var запрос = СформироватьЗапросHttp();
        using var ответ  = await СформироватьОтвет();

        return JsonSerializer.Serialize(ответ, jsonSerializerOptions);


        HttpRequestMessage СформироватьЗапросHttp() {
            var тело = телоJson is null
                           ? null
                           : new StringContent(телоJson, Encoding.UTF8, MediaTypeNames.Application.Json);

            var запросHttp = new HttpRequestMessage { Method = метод, RequestUri = new Uri(адрес), Content = тело };

            for (var i = 0; i < заголовкиСтр.GetLength(0); i++) {
                var название = заголовкиСтр[i, 0];
                var значения = Enumerable
                              .Range(1, заголовкиСтр.GetLength(1) - 1)
                               // ReSharper disable once AccessToModifiedClosure
                              .Select(j => заголовкиСтр[i, j])
                              .Select(стр => string.IsNullOrWhiteSpace(стр) ? null : стр.Trim())
                              .Where(v => !string.IsNullOrWhiteSpace(v));
                запросHttp.Headers.Add(название, значения);
            }

            if (string.IsNullOrWhiteSpace(заголовкиJson)) return запросHttp;

            var заголовкиСписок =
                JsonSerializer.Deserialize<List<ОбъектJson>>(заголовкиJson, jsonSerializerOptions);
            var допЗаголовки = from заголовокObj in заголовкиСписок
                               from заголовок in заголовокObj
                               let имя = заголовок.Key
                               let парамм = заголовок.Value.ValueKind == JsonValueKind.Array
                                                ? заголовок.Value.EnumerateArray()
                                                : new[] { заголовок.Value }.AsEnumerable()
                               let параммStr = парамм
                                  .Select(п => п.ValueKind == JsonValueKind.String
                                                   ? п.GetString()
                                                   : п.GetRawText())
                               select (имя, параммStr);

            foreach (var (имя, параммStr) in допЗаголовки) запросHttp.Headers.Add(имя, параммStr);

            return запросHttp;
        }

        JsonDocument? ПопробоватьРазобратьJsonТекст(string jsonТекст) {
            try {
                return JsonDocument.Parse(jsonТекст, new JsonDocumentOptions { AllowTrailingCommas = true });
            }
            catch {
                return null;
            }
        }

        async Task<ОтветHttp> СформироватьОтвет() {
            using var ответHttpСообщение = await httpКлиент.SendAsync(запрос, ct).ConfigureAwait(false);
            var содержаниеStr = ответHttpСообщение is { IsSuccessStatusCode: true }
                                    ? await ответHttpСообщение.Content.ReadAsStringAsync(ct).ConfigureAwait(false)
                                    : null;

            var json = ПопробоватьРазобратьJsonТекст(содержаниеStr ?? "");

            return new ОтветHttp(ответHttpСообщение.StatusCode,
                                 DateTime.Now,
                                 (object?)json ?? содержаниеStr,
                                 ответHttpСообщение.Headers,
                                 ответHttpСообщение.Content.Headers,
                                 ответHttpСообщение.Headers.TryGetValues("Set-Cookie", out var куки) ? куки : null);
        }
    }

    public record ОтветHttp(
        [property: JsonPropertyName("status")] HttpStatusCode Статус,
        [property: JsonPropertyName("date")] DateTime ДатаЗапроса,
        [property: JsonPropertyName("result")] object? Содержимое,
        [property: JsonPropertyName("header")] HttpResponseHeaders? Заголовки,
        [property: JsonPropertyName("header2")] HttpContentHeaders? Заголовки2,
        [property: JsonPropertyName("cookie")] IEnumerable<string>? Куки) : IDisposable {
        private bool _утилизированЛи;

        public void Dispose() {
            if (_утилизированЛи) return;
            if (Содержимое is JsonDocument json) json.Dispose();
            _утилизированЛи = true;
        }
    }
}