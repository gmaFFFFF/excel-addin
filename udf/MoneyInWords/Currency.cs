using System.Globalization;
using System.Text;
using static DecimalMath.DecimalEx;

namespace gmafffff.excel.udf.ДеньгиПрописью;

public sealed partial class ДеньгиПрописью {
    public class Валюта : IFormattable {
        private ЧислоЦелоеКлассы _дробнаяЧасть;
        private ЧислоЦелоеКлассы _целаяЧасть;
        private decimal _число;

        public Валюта(ИСправочникЧисел справочник, decimal число, ВалютаIso валютаIso) {
            ArgumentNullException.ThrowIfNull(справочник);

            Справочник = справочник;
            ВалютаIso = валютаIso;
            Число = число;
        }

        /// <summary>
        ///     Справочник наименований на определённом языке
        /// </summary>
        public ИСправочникЧисел Справочник { get; set; }

        /// <summary>
        ///     Значение в валюте
        /// </summary>
        public decimal Число {
            get => _число;
            set {
                _число = value;
                _целаяЧасть = new ЧислоЦелоеКлассы(ЦелаяЧасть, Справочник);
                _дробнаяЧасть = new ЧислоЦелоеКлассы(ДробнаяЧасть, Справочник);
            }
        }

        public ВалютаIso ВалютаIso { get; set; }

        /// <summary>
        ///     Целая часть <see cref="Число" />
        /// </summary>
        public decimal ЦелаяЧасть => Floor(Число);

        /// <summary>
        ///     Дробная часть <see cref="Число" />
        /// </summary>
        public decimal ДробнаяЧасть =>
            Floor((Число - ЦелаяЧасть) * Pow(10, Справочник.ДробнаяЧастьЧислоЗнаков[ВалютаIso]));

        /// <summary>
        /// </summary>
        /// <param name="format">
        ///     определяет строку в соответствии с которой форматируется число
        ///     Допустимые элементы формата:
        ///     ч[n] - деньги в числовом формате, n - число знаков после запятой
        ///     б[n][т[з]] - целая часть валюты, т - текстом, з - с заглавной буквы, n - ширина
        ///     д[n][т[з]] - дробная часть валюты, т - текстом, з - с заглавной буквы, n - ширина
        ///     р[с] - валюта базовая, с - сокращенная
        ///     к[c] - валюта дробная, с - сокращенная
        /// </param>
        /// <returns></returns>
        public string ToString(string? format, IFormatProvider? formatProvider) {
            const char пробел_между_разрядами = ' ';
            if (string.IsNullOrEmpty(format) || format == "G") format = "ч2 рс (бтз р д2 к)";

            formatProvider ??= CultureInfo.CurrentCulture;
            format = format.ToLower();

            StringBuilder итог = new();

            Queue<char> символы = new(format.ToCharArray());

            while (символы.Count > 0) {
                var тек = символы.Dequeue() switch {
                    (char)СимволыФормата.ВалютаЦеликом => ПреобразуйВалютуВСтроку(),
                    (char)СимволыФормата.ВалютаБазовая => ФорматируйВалютуВСтроку(_целаяЧасть),
                    (char)СимволыФормата.ВалютаДробная => ФорматируйВалютуВСтроку(_дробнаяЧасть, '0'),
                    (char)СимволыФормата.ВалютаОбознБазовая => ФорматируйОбознВалютыБ(),
                    (char)СимволыФормата.ВалютаОбознДробная => ФорматируйОбознВалютыД(),
                    var sym => sym.ToString()
                };
                итог.Append(тек);
            }

            return итог.ToString();

            string ФорматируйОбознВалютыБ() {
                if (!символы.TryPeek(out var след) || след != (char)СимволыФормата.ВалютаОбознСокр)
                    return ДайПодписьВалютыБазовой();
                символы.Dequeue();
                return ДайПодписьВалютыБазовой(true);
            }

            string ФорматируйОбознВалютыД() {
                if (!символы.TryPeek(out var след) || след != (char)СимволыФормата.ВалютаОбознСокр)
                    return ДайПодписьВалютыДробной();
                символы.Dequeue();
                return ДайПодписьВалютыДробной(true);
            }

            string ФорматируйВалютуВСтроку(ЧислоЦелоеКлассы число, char заполнитель = ' ') {
                return ФорматируйПервуюБукву(ФорматируйВЧислоИлиТекст(число));


                string ФорматируйПервуюБукву(string текст) {
                    if (!символы.TryPeek(out var след) || след != (char)СимволыФормата.ЗаглавняБуква)
                        return текст;

                    символы.Dequeue();
                    return ПреобрПервуюБуквуВЗаглавную(текст);
                }

                string ФорматируйВЧислоИлиТекст(ЧислоЦелоеКлассы число) {
                    if (!символы.TryPeek(out var след) || след != (char)СимволыФормата.ЧислоТекстом)
                        return ПреобразуйЧислоВСтроку(число);

                    символы.Dequeue();
                    return число.ПреобразуйВТекст();
                }

                string ПреобразуйЧислоВСтроку(ЧислоЦелоеКлассы число) {
                    var шир = ВытолкниЧисло() ?? 0;
                    var формат = ((decimal)число).ToString("N0");
                    var формат_пробел = формат.Replace(' ', пробел_между_разрядами);
                    var формат_пробел_шир = формат_пробел.PadLeft(шир, заполнитель);
                    return формат_пробел_шир;
                }

                string ПреобрПервуюБуквуВЗаглавную(string стр) {
                    return string.Concat(стр[0].ToString().ToUpper(), стр.AsSpan(1));
                }
            }

            string ПреобразуйВалютуВСтроку() {
                var фрмт = (NumberFormatInfo?)formatProvider.GetFormat(typeof(NumberFormatInfo));
                var округ = ВытолкниЧисло() ?? фрмт?.CurrencyDecimalDigits ?? 2;

                return _число.ToString($"N{округ}", formatProvider).Replace(' ', пробел_между_разрядами);
            }

            int? ВытолкниЧисло() {
                var число_текст = new string(ВытолкниЧисловыеСимволы().ToArray());
                return int.TryParse(число_текст, out var рез)
                    ? рез
                    : null;
            }

            IEnumerable<char> ВытолкниЧисловыеСимволы() {
                while (символы.Count > 0)
                    if (char.IsNumber(символы.Peek()))
                        yield return символы.Dequeue();
                    else yield break;
            }
        }

        public string ДайЦелуюЧастьСловами(bool родМужскойЛи = true) {
            return _целаяЧасть.ПреобразуйВТекст(родМужскойЛи);
        }

        public string ДайДробнуюЧастьСловами(bool родМужскойЛи = true) {
            return _дробнаяЧасть.ПреобразуйВТекст(родМужскойЛи);
        }

        public string ДайПодписьВалютыБазовой(bool сокращенноеЛи = false) {
            return сокращенноеЛи
                ? Справочник.ДайНазваниеБазовойВалютыСокр(_целаяЧасть, ВалютаIso)
                : Справочник.ДайНазваниеБазовойВалютыПолное(_целаяЧасть, ВалютаIso);
        }

        public string ДайПодписьВалютыДробной(bool сокращенноеЛи = false) {
            return сокращенноеЛи
                ? Справочник.ДайНазваниеДробнойВалютыСокр(_дробнаяЧасть, ВалютаIso)
                : Справочник.ДайНазваниеДробнойВалютыПолное(_дробнаяЧасть, ВалютаIso);
        }


        public override string ToString() {
            return ToString("G", CultureInfo.CurrentCulture);
        }

        public string ToString(string format) {
            return ToString(format, CultureInfo.CurrentCulture);
        }

        private enum СимволыФормата {
            ВалютаЦеликом = 'ч',
            ВалютаБазовая = 'б',
            ВалютаДробная = 'д',
            ВалютаОбознБазовая = 'р',
            ВалютаОбознДробная = 'к',
            ВалютаОбознСокр = 'с',
            ЧислоТекстом = 'т',
            ЗаглавняБуква = 'з'
        }
    }
}