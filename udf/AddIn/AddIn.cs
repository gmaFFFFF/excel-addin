using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;

namespace gmafffff.excel.udf.AddIn;

public sealed class AddIn : IExcelAddIn {
    private static ParameterConversionConfiguration автоПриведениеТиповКонфиг
        => new ParameterConversionConfiguration()
            // Добавляет поддержку параметров string[] (вместо этого принимается object[]).
            // Использует служебный класс TypeConversion, определенный в ExcelDna.Registration, 
            // преобразование выполняется Excel.
            .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())
            // Добавляет поддержку параметров string[,] (вместо этого принимается object[,]).
            .AddParameterConversion((object[,] arr) => Массив2dОбъектовВМассив2dСтрок(arr))
            // Пара очень общих преобразований для типов Enum
            .AddReturnConversion((Enum value) => value.ToString(), true)
            .AddParameterConversion(ParameterConversions.GetEnumStringConversion());

    public void AutoOpen() {
        ExcelIntegration.RegisterUnhandledExceptionHandler(ex => $"!!! Ошибка: {ex}");
        РегистрироватьФункции();

        IntelliSenseServer.Install(); // Важен порядок вызова после регистрации функций, иначе нужен Refresh
    }

    public void AutoClose() {
        IntelliSenseServer.Uninstall();
    }

    private static void РегистрироватьФункции() {
        ExcelRegistration.GetExcelFunctions()
            .ProcessMapArrayFunctions()
            .ProcessParameterConversions(автоПриведениеТиповКонфиг)
            // Преимущество встроенной поддержки асинхронных функций (nativeAsyncIfAvailable):
            //   — пока функция вычисляется в ячейке будет отображаться #ВЫЧИСЛ! вместо #Н/Д
            // Недостатки встроенной поддержки асинхронных функций:
            //   — требуется два раза нажать Enter, чтобы продолжить ввод даных;
            //   — субъективно работает медленнее.
            .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
            .ProcessParamsRegistrations()
            .RegisterFunctions();
    }


    private static string[,] Массив2dОбъектовВМассив2dСтрок(object[,] массив) {
        var массивН = new string[массив.GetLength(0), массив.GetLength(1)];
        for (var i = 0; i < массив.GetLength(0); i++)
        for (var j = 0; j < массив.GetLength(1); j++)
            массивН[i, j] = TypeConversion.ConvertToString(массив[i, j]);

        return массивН;
    }
}