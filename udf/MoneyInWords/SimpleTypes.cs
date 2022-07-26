namespace gmafffff.excel.udf.ДеньгиПрописью;

public sealed partial class ДеньгиПрописью {
    // Код валюты по ISO 4217
    public enum ВалютаIso : ushort {
        Rub = 643
    }

    /// <summary>
    ///     Разряд класса числа
    /// </summary>
    public enum Разряд : byte {
        Единицы = 1,
        Десятки = 10,
        Сотни = 100
    }

    public enum Степень10 : byte {
        Едн = 0,
        Тыс = 3,
        Млн = 6,
        Млрд = 9,
        Трлн = 12,
        Квдлн = 15,
        Квнлн = 18,
        Сктлн = 21,
        Сптлн = 24,
        Октилн = 27,
        Нонилн = 30,
        Децилн = 33,
        Ундец = 36,
        Дуодец = 39,
        Тредец = 42,
        Кваттуор = 45,
        Квиндец = 48,
        Сексдец = 51,
        Септдец = 54,
        Октодец = 57,
        Новемдец = 60,
        Вигинт = 63,
        Унвигинт = 66,
        Дуовигинт = 69,
        Тревигинт = 72,
        Кваттуорвигинт = 75,
        Квинвигинт = 78,
        Сексвигинт = 81,
        Септенвигинт = 84,
        Октовигинт = 87,
        Новемвигинт = 90,
        Тригинт = 93,
        Унтригинт = 96,
        Дуотригинт = 99,
        Третригинт = 102,
        Кваттуортригинт = 105,
        Квинтригинт = 108,
        Секстригинт = 111,
        Септентригинт = 114,
        Октотригинт = 117,
        Новемтригинт = 120,
        Квадрагинт = 123
    }

    /// <summary>
    ///     Названия цифр у разрядов числа
    /// </summary>
    /// <param name="Разряд">
    ///     1   - единицы
    ///     10  - десятки
    ///     100 - сотни
    /// </param>
    /// <param name="Названия">Название разрядов числа</param>
    public record РазрядЧислаНазванияЦифр(Разряд Разряд, string[] Названия);

    /// <summary>
    ///     Склонение числа
    /// </summary>
    /// <param name="Цифры">Цифры</param>
    /// <param name="Склонение">Склонение</param>
    public record СклонениеЧисла(ushort Цифры, string Склонение);

    /// <summary>
    ///     Склонения классов чисел
    /// </summary>
    /// <param name="Степень10">Класс числа, выражающийся через степень 10 первого разряда</param>
    /// <param name="Склонения">Склонения класса</param>
    /// <param name="Стандартное">Склонение класса применяемое во всех остальных случаях</param>
    public record КлассЧислаНазваниеСклонения(Степень10 Степень10, СклонениеЧисла[] Склонения, string Стандартное);

    /// <summary>
    ///     Склоения обозначений валют
    /// </summary>
    /// <param name="Склонения">Склонения валюты</param>
    /// <param name="Стандартное">Склонение обозначения валюты применяемое во всех остальных случаях</param>
    public record ВалютаСклонения(СклонениеЧисла[] Склонения, string Стандартное);
}