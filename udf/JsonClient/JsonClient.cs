using System.Text.Json;
using JsonCons.JmesPath;
using JsonCons.JsonPath;

namespace gmafffff.excel.udf.Json������;

public sealed class Json������ {
    private static readonly JsonDocumentOptions Json�������� = new() { AllowTrailingCommas = true };

    /// <summary>
    ///     ������� ������� �� JSONPath
    /// </summary>
    /// <param name="json�����">����� � ������� JSON</param>
    /// <param name="������">
    ///     ������ � ������� JSONPath.
    ///     ��������� � ������� JSONPath ��
    ///     <a href="https://danielaparker.github.io/JsonCons.Net/articles/JsonPath/JsonConsJsonPath.html">������</a>
    /// </param>
    /// <returns>�������� ���������� �������� ��� ������ JSON</returns>
    public static object JsonPath�����(string json�����, string ������) {
        if (string.IsNullOrWhiteSpace(json�����) || string.IsNullOrWhiteSpace(������))
            return "";

        var ��������������� = JsonSelectorOptions.Default;
        ���������������.ExecutionMode = PathExecutionMode.Parallel;

        using var json = JsonDocument.Parse(json�����, Json��������);

        var ������� = JsonSelector.Select(json.RootElement, ������, ���������������);

        return �������������Json�������(�������);
    }

    public static object JmesPath������(string json�����, string ������) {
        if (string.IsNullOrWhiteSpace(json�����) || string.IsNullOrWhiteSpace(������))
            return "";

        var ��������������� = JsonSelectorOptions.Default;
        ���������������.ExecutionMode = PathExecutionMode.Parallel;


        using var json = JsonDocument.Parse(json�����, Json��������);

        var ������� = JsonTransformer.Transform(json.RootElement, ������);

        var ���� = ������� is not null
            ? new[] { �������.RootElement }
            : Array.Empty<JsonElement>();

        return �������������Json�������(����);
    }

    private static object �������������Json�������(IList<JsonElement> �������) {
        var ��������������� = new JsonSerializerOptions { WriteIndented = true, AllowTrailingCommas = true };

        return �������.Count switch {
            0 => "",
            1 => ���������������������������Json(�������[0], ���������������),
            _ => JsonSerializer.Serialize(�������, ���������������)
        };
    }

    private static object ���������������������������Json(JsonElement �������, JsonSerializerOptions? ���������������) {
        return �������.ValueKind switch {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => "",
            JsonValueKind.Number => ���������������������������Json��������(�������, ���������������),
            JsonValueKind.String => �������.TryGetDateTime(out var ����)
                ? ����
                : �������.GetString(),
            _ => JsonSerializer.Serialize(�������, ���������������)
        } ?? "";
    }

    private static object ���������������������������Json��������(JsonElement �������,
        JsonSerializerOptions? ���������������) {
        if (�������.TryGetInt64(out var �����))
            return �����;
        if (�������.TryGetDecimal(out var ���))
            return ���;
        if (�������.TryGetDouble(out var ���))
            return ���;

        return JsonSerializer.Serialize(�������, ���������������);
    }
}