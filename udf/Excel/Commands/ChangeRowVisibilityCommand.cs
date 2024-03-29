using gmafffff.excel.udf.Excel.Сетка;

namespace gmafffff.excel.udf.Excel.Команды;

public sealed record ИзмениВидимостьРядаКоманда(
    РядСетки Ряд,
    bool Показать,
    double? Размер = null,
    bool Авторазмер = false) : ИExcelКоманда {
    public void Выполнить() {
        if (Показать) Ряд.Показать(Размер, Авторазмер);
        Ряд.Скрыть();
    }

    public static IEnumerable<ИзмениВидимостьРядаКомандаПакет> Упаковать(IEnumerable<ИExcelКоманда> команды) {
        var командыСхожие = команды
                           .OfType<ИзмениВидимостьРядаКоманда>()
                           .GroupBy(e => (e.Показать, e.Размер, e.Авторазмер));

        // Упаковываем команды в один пакет
        var пакетКоманд =
            from группа in командыСхожие
            select new ИзмениВидимостьРядаКомандаПакет(from команда in группа select команда.Ряд,
                                                       группа.Key.Показать,
                                                       группа.Key.Размер,
                                                       группа.Key.Авторазмер);
        return пакетКоманд;
    }
}

public record ИзмениВидимостьРядаКомандаПакет(
    IEnumerable<РядСетки> Ряды,
    bool Показать,
    double? Размер = null,
    bool Авторазмер = false) : ИExcelКоманда {
    public void Выполнить() {
        var нужноИсполнить = Ряды
                             // Опускаем команды, которые выполнять не требуется
                            .Where(ряд => Показать != ряд.ВидимЛи())
                             // Из двух команд, касающихся одного и того же ряда, удаляем ту, которая поступила раньше
                            .Select((ряд, i) => (ряд, i))
                            .Reverse()
                            .DistinctBy(e => (e.ряд.GetType(), e.ряд))
                            .OrderBy(e => e.i)
                            .Select(e => e.ряд)
                             // Делим на две группы: строки и столбцы
                            .GroupBy(р => р.GetType());

        foreach (var ряды in нужноИсполнить) {
            РядСетки.ИзменитьВидимость(ряды.Select(р => р.GetAsExcelReference()),
                                       ряды.Key,
                                       Показать ? Размер : -1,
                                       Авторазмер);
        }
    }
}