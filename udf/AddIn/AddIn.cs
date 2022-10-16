using ExcelDna.Integration;
using ExcelDna.Registration;

namespace gmafffff.excel.udf.AddIn;

public sealed class AddIn : IExcelAddIn {
    private static ParameterConversionConfiguration �������������������������
        => new ParameterConversionConfiguration()
            // ��������� ��������� ���������� string[] (������ ����� ����������� object[]).
            // ���������� ��������� ����� TypeConversion, ������������ � ExcelDna.Registration, 
            // �������������� ����������� Excel.
            .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())
            // ��������� ��������� ���������� string[,] (������ ����� ����������� object[,]).
            .AddParameterConversion((object[,] arr) => ������2d���������������2d�����(arr))
            // ���� ����� ����� �������������� ��� ����� Enum
            .AddReturnConversion((Enum value) => value.ToString(), true)
            .AddParameterConversion(ParameterConversions.GetEnumStringConversion());

    public void AutoOpen() {
        ExcelIntegration.RegisterUnhandledExceptionHandler(ex => $"!!! ������: {ex}");
        ���������������������();
    }

    public void AutoClose() { }

    private static void ���������������������() {
        ExcelRegistration.GetExcelFunctions()
            .ProcessMapArrayFunctions()
            .ProcessParameterConversions(�������������������������)
            .ProcessAsyncRegistrations(true)
            .ProcessParamsRegistrations()
            .RegisterFunctions();
    }


    private static string[,] ������2d���������������2d�����(object[,] ������) {
        var ������� = new string[������.GetLength(0), ������.GetLength(1)];
        for (var i = 0; i < ������.GetLength(0); i++)
        for (var j = 0; j < ������.GetLength(1); j++)
            �������[i, j] = TypeConversion.ConvertToString(������[i, j]);

        return �������;
    }
}