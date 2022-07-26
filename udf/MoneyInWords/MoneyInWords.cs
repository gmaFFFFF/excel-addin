namespace gmafffff.excel.udf.ДеньгиПрописью;

/// <summary>
/// </summary>
/// <remarks> Вдохновение черпал здесь: http://www.num2word.ru</remarks>
public sealed partial class ДеньгиПрописью {
    /// <summary>
    /// </summary>
    /// <param name="число"></param>
    /// <param name="формат">
    ///     Допустимые элементы формата:
    ///     ч[n] - деньги в числовом формате, n - число знаков после запятой
    ///     б[n][т[з]] - целая часть валюты, т - текстом, з - с заглавной буквы, n - ширина
    ///     д[n][т[з]] - дробная часть валюты, т - текстом, з - с заглавной буквы, n - ширина
    ///     р[с] - валюта базовая, с - сокращенная
    ///     к[c] - валюта дробная, с - сокращенная
    /// </param>
    /// <returns></returns>
    public static string РублиПрописью(double число, string формат = "") {
        try {
            return new Валюта(new СправочникЧиселРус(), (decimal)число, ВалютаIso.Rub).ToString(формат);
        }
        catch (OverflowException e) {
            return "Такое большое число не поддерживается";
        }
    }
}