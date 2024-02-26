using System.Text;
using static DecimalMath.DecimalEx;

namespace gmafffff.excel.udf.ДеньгиПрописью;

public sealed partial class ДеньгиПрописью {
    /// <summary>
    ///     Цифры в записи многозначных чисел разбивают справа налево на группы по три цифры в каждой.
    ///     <br />Эти группы называют классами.
    ///     <br />В каждом классе цифры справа налево обозначают единицы, десятки и сотни этого класса:
    /// </summary>
    public class ЧислоЦелоеКлассы {
        /// <summary>
        ///     Справочник наименований цифр на определённом языке
        /// </summary>
        private readonly ИСправочникЧисел _справочник;

        private readonly decimal _число;

        private Dictionary<Степень10, ЧислоЦелоеКласс> _числоПоКлассам;

        public ЧислоЦелоеКлассы(decimal число, ИСправочникЧисел справочник) {
            if (число < 0m)
                throw new ArgumentException($"Число должен быть больше или равно 0, а передано {число}",
                                            nameof(число));

            _справочник = справочник;
            _число      = decimal.Ceiling(число);
            РазбитьНаКлассы(_число);
        }

        public ЧислоЦелоеКласс this[Степень10 степень] =>
            _числоПоКлассам.TryGetValue(степень, out var val)
                ? val
                : new ЧислоЦелоеКласс(0, _справочник);

        public ЧислоЦелоеКласс ПоследнийЗначимыйКласс
            => _числоПоКлассам.Values
                              .LastOrDefault(к => к.ЗначимыйЛи)
            ?? new ЧислоЦелоеКласс(0, _справочник);

        private void РазбитьНаКлассы(decimal число) {
            var разрядов = число != 0 ? Math.Truncate(Log10(число)) + 1 : 1;
            var классов  = (byte)Math.Ceiling(разрядов / 3);

            _числоПоКлассам = new Dictionary<Степень10, ЧислоЦелоеКласс>(классов);
            foreach (var i in Enumerable.Range(1, классов).Reverse()) {
                var классГрВ = Pow(10, i * 3);
                var классГрН = Pow(10, (i - 1) * 3);
                var степень  = (Степень10)((i - 1) * 3);

                var класс = new ЧислоЦелоеКласс((ushort)(число % классГрВ / классГрН), _справочник);
                if (класс.ЗначимыйЛи) _числоПоКлассам[степень] = класс;
            }
        }

        public string ПреобразуйВТекст(bool родМужскойЛи = true) {
            StringBuilder sb = new();

            foreach (var (степень, числоВКлассе) in _числоПоКлассам) {
                var названиеКласса = _справочник.ДайНазваниеКласса(степень, числоВКлассе);

                var род = (_справочник.КлассТребуетЖенРодЛи(степень, числоВКлассе), степень) switch {
                    (true, _)          => false,
                    (_, Степень10.Едн) => родМужскойЛи,
                    _                  => true
                };

                var часть = string.IsNullOrEmpty(названиеКласса)
                                ? числоВКлассе.ПреобразуйВТекст(род)
                                : string.Join(' ', числоВКлассе.ПреобразуйВТекст(род), названиеКласса);

                if (sb is not { Length: 0 }) sb.Append(' ');
                sb.Append(часть);
            }

            return sb.Length == 0
                       ? _справочник.ДайНазваниеИменногоЧисла(0, родМужскойЛи)
                       : sb.ToString();
        }

        public override string ToString() { return ПреобразуйВТекст(); }

        public static implicit operator decimal(ЧислоЦелоеКлассы чцк) { return чцк._число; }
    }
}