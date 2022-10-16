using System.Net;
using System.Net.Http.Headers;
using System.Net.Mime;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace gmafffff.excel.udf.HttpКлиент;

public sealed class HttpКлиент {
    private static readonly HttpClient Клиент;

    static HttpКлиент() {
        Клиент = new HttpClient();
    }

    public static async Task<string?> GetАсинх(string адрес, string[,]? заголовкиСтр = null,
        CancellationToken ct = default) {
        return await ПослатьАсинх(HttpMethod.Get, адрес, заголовкиСтр ?? new string[0, 0], null, ct)
            .ConfigureAwait(false);
    }

    public static async Task<string?> PostАсинх(string адрес, string[,]? заголовкиСтр, string? телоJson,
        CancellationToken ct = default) {
        return await ПослатьАсинх(HttpMethod.Post, адрес, заголовкиСтр ?? new string[0, 0], телоJson, ct)
            .ConfigureAwait(false);
    }

    public static async Task<string?> ПослатьАсинх(HttpMethod метод, string адрес,
        string[,] заголовкиСтр, string? телоJson,
        CancellationToken ct = default) {
        if (string.IsNullOrWhiteSpace(адрес))
            return null;

        var jsonSerializerOptions = new JsonSerializerOptions {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        телоJson = string.IsNullOrWhiteSpace(телоJson) ? null : телоJson.Trim();
        адрес = адрес.Trim();

        using var запрос = СформироватьЗапросHttp();
        using var ответHttpСообщение = await Клиент.SendAsync(запрос, ct).ConfigureAwait(false);
        using var ответ = await СформироватьОтвет();

        return JsonSerializer.Serialize(ответ, jsonSerializerOptions);


        HttpRequestMessage СформироватьЗапросHttp() {
            var тело = телоJson is null
                ? null
                : new StringContent(телоJson, Encoding.UTF8, MediaTypeNames.Application.Json);

            var запросHttp = new HttpRequestMessage {
                Method = метод,
                RequestUri = new Uri(адрес),
                Content = тело
            };

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

            return запросHttp;
        }

        JsonDocument? ПопробоватьРазобратьJsonТекст(string jsonТекст) {
            try {
                return JsonDocument.Parse(jsonТекст, new JsonDocumentOptions() { AllowTrailingCommas = true });
            }
            catch {
                return null;
            }
        }

        async Task<ОтветHttp> СформироватьОтвет() {
            var содержаниеStr = ответHttpСообщение is { IsSuccessStatusCode: true }
                ? await ответHttpСообщение.Content.ReadAsStringAsync(ct).ConfigureAwait(false)
                : null;

            var json = ПопробоватьРазобратьJsonТекст(содержаниеStr ?? "");

            return new ОтветHttp(ответHttpСообщение.StatusCode,
                ответHttpСообщение.Headers,
                ответHttpСообщение.Content.Headers,
                ответHttpСообщение.Headers.TryGetValues("Set-Cookie", out var куки)
                    ? куки
                    : null,
                (object?)json ?? содержаниеStr,
                DateTime.Now
            );
        }
    }

    public record ОтветHttp(HttpStatusCode Статус, HttpResponseHeaders? Заголовки, HttpContentHeaders? Заголовки2,
        IEnumerable<string>? Куки, object? Содержимое, DateTime ДатаЗапроса) : IDisposable {
        private bool _утилизированЛи = false;

        public void Dispose() {
            if (_утилизированЛи)
                return;
            if (Содержимое is JsonDocument json)
                json.Dispose();
            _утилизированЛи = true;
        }
    }
}