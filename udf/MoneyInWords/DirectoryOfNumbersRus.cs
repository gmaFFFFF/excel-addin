using System.Collections.Immutable;

namespace gmafffff.excel.udf.ДеньгиПрописью;

public sealed partial class ДеньгиПрописью {
    public class СправочникЧиселРус : ИСправочникЧисел {
        protected static readonly КлассЧислаНазваниеСклонения[] КлассЧислаНазваниеСклонение = {
            new(Степень10.Едн,
                Array.Empty<СклонениеЧисла>(),
                ""),
            new(Степень10.Тыс,
                new[] {
                    new СклонениеЧисла(0, "тысяч"), new СклонениеЧисла(1, "тысяча"), new СклонениеЧисла(2, "тысячи"),
                    new СклонениеЧисла(3, "тысячи"), new СклонениеЧисла(4, "тысячи")
                },
                "тысяч"),
            new(Степень10.Млн,
                new[] {
                    new СклонениеЧисла(0, "миллионов"), new СклонениеЧисла(1, "миллион"),
                    new СклонениеЧисла(2, "миллиона"), new СклонениеЧисла(3, "миллиона"),
                    new СклонениеЧисла(4, "миллиона")
                },
                "миллионов"),
            new(Степень10.Млрд,
                new[] {
                    new СклонениеЧисла(0, "миллиардов"), new СклонениеЧисла(1, "миллиард"),
                    new СклонениеЧисла(2, "миллиарда"), new СклонениеЧисла(3, "миллиарда"),
                    new СклонениеЧисла(4, "миллиарда")
                },
                "миллиардов"),
            new(Степень10.Трлн,
                new[] {
                    new СклонениеЧисла(0, "триллионов"), new СклонениеЧисла(1, "триллион"),
                    new СклонениеЧисла(2, "триллиона"), new СклонениеЧисла(3, "триллиона"),
                    new СклонениеЧисла(4, "триллиона")
                },
                "триллионов"),
            new(Степень10.Квдлн,
                new[] {
                    new СклонениеЧисла(0, "квадриллионов"), new СклонениеЧисла(1, "квадриллион"),
                    new СклонениеЧисла(2, "квадриллиона"), new СклонениеЧисла(3, "квадриллиона"),
                    new СклонениеЧисла(4, "квадриллиона")
                },
                "квадриллионов"),
            new(Степень10.Квнлн,
                new[] {
                    new СклонениеЧисла(0, "квинтиллионов"), new СклонениеЧисла(1, "квинтиллион"),
                    new СклонениеЧисла(2, "квинтиллиона"), new СклонениеЧисла(3, "квинтиллиона"),
                    new СклонениеЧисла(4, "квинтиллиона")
                },
                "квинтиллионов"),
            new(Степень10.Сктлн,
                new[] {
                    new СклонениеЧисла(0, "секстиллионов"), new СклонениеЧисла(1, "секстиллион"),
                    new СклонениеЧисла(2, "секстиллиона"), new СклонениеЧисла(3, "секстиллиона"),
                    new СклонениеЧисла(4, "секстиллиона")
                },
                "секстиллионов"),
            new(Степень10.Сптлн,
                new[] {
                    new СклонениеЧисла(0, "септиллионов"), new СклонениеЧисла(1, "септиллион"),
                    new СклонениеЧисла(2, "септиллиона"), new СклонениеЧисла(3, "септиллиона"),
                    new СклонениеЧисла(4, "септиллиона")
                },
                "септиллионов"),
            new(Степень10.Октилн,
                new[] {
                    new СклонениеЧисла(0, "октиллионов"), new СклонениеЧисла(1, "октиллион"),
                    new СклонениеЧисла(2, "октиллиона"), new СклонениеЧисла(3, "октиллиона"),
                    new СклонениеЧисла(4, "октиллиона")
                },
                "октиллионов"),
            new(Степень10.Нонилн,
                new[] {
                    new СклонениеЧисла(0, "нониллионов"), new СклонениеЧисла(1, "нониллион"),
                    new СклонениеЧисла(2, "нониллиона"), new СклонениеЧисла(3, "нониллиона"),
                    new СклонениеЧисла(4, "нониллиона")
                },
                "нониллионов"),
            new(Степень10.Децилн,
                new[] {
                    new СклонениеЧисла(0, "дециллионов"), new СклонениеЧисла(1, "дециллион"),
                    new СклонениеЧисла(2, "дециллиона"), new СклонениеЧисла(3, "дециллиона"),
                    new СклонениеЧисла(4, "дециллиона")
                },
                "дециллионов"),
            new(Степень10.Ундец,
                new[] {
                    new СклонениеЧисла(0, "ундециллионов"), new СклонениеЧисла(1, "ундециллион"),
                    new СклонениеЧисла(2, "ундециллиона"), new СклонениеЧисла(3, "ундециллиона"),
                    new СклонениеЧисла(4, "ундециллиона")
                },
                "ундециллионов"),
            new(Степень10.Дуодец,
                new[] {
                    new СклонениеЧисла(0, "дуодециллионов"), new СклонениеЧисла(1, "дуодециллион"),
                    new СклонениеЧисла(2, "дуодециллиона"), new СклонениеЧисла(3, "дуодециллиона"),
                    new СклонениеЧисла(4, "дуодециллиона")
                },
                "дуодециллионов"),
            new(Степень10.Тредец,
                new[] {
                    new СклонениеЧисла(0, "тредециллионов"), new СклонениеЧисла(1, "тредециллион"),
                    new СклонениеЧисла(2, "тредециллиона"), new СклонениеЧисла(3, "тредециллиона"),
                    new СклонениеЧисла(4, "тредециллиона")
                },
                "тредециллионов"),
            new(Степень10.Кваттуор,
                new[] {
                    new СклонениеЧисла(0, "кваттуордециллионов"), new СклонениеЧисла(1, "кваттуордециллион"),
                    new СклонениеЧисла(2, "кваттуордециллиона"), new СклонениеЧисла(3, "кваттуордециллиона"),
                    new СклонениеЧисла(4, "кваттуордециллиона")
                },
                "кваттуордециллионов"),
            new(Степень10.Квиндец,
                new[] {
                    new СклонениеЧисла(0, "квиндециллионов"), new СклонениеЧисла(1, "квиндециллион"),
                    new СклонениеЧисла(2, "квиндециллиона"), new СклонениеЧисла(3, "квиндециллиона"),
                    new СклонениеЧисла(4, "квиндециллиона")
                },
                "квиндециллионов"),
            new(Степень10.Сексдец,
                new[] {
                    new СклонениеЧисла(0, "сексдециллионов"), new СклонениеЧисла(1, "сексдециллион"),
                    new СклонениеЧисла(2, "сексдециллиона"), new СклонениеЧисла(3, "сексдециллиона"),
                    new СклонениеЧисла(4, "сексдециллиона")
                },
                "сексдециллионов"),
            new(Степень10.Септдец,
                new[] {
                    new СклонениеЧисла(0, "септдециллионов"), new СклонениеЧисла(1, "септдециллион"),
                    new СклонениеЧисла(2, "септдециллиона"), new СклонениеЧисла(3, "септдециллиона"),
                    new СклонениеЧисла(4, "септдециллиона")
                },
                "септдециллионов"),
            new(Степень10.Октодец,
                new[] {
                    new СклонениеЧисла(0, "октодециллионов"), new СклонениеЧисла(1, "октодециллион"),
                    new СклонениеЧисла(2, "октодециллиона"), new СклонениеЧисла(3, "октодециллиона"),
                    new СклонениеЧисла(4, "октодециллиона")
                },
                "октодециллионов"),
            new(Степень10.Новемдец,
                new[] {
                    new СклонениеЧисла(0, "новемдециллионов"), new СклонениеЧисла(1, "новемдециллион"),
                    new СклонениеЧисла(2, "новемдециллиона"), new СклонениеЧисла(3, "новемдециллиона"),
                    new СклонениеЧисла(4, "новемдециллиона")
                },
                "новемдециллионов"),
            new(Степень10.Вигинт,
                new[] {
                    new СклонениеЧисла(0, "вигинтиллионов"), new СклонениеЧисла(1, "вигинтиллион"),
                    new СклонениеЧисла(2, "вигинтиллиона"), new СклонениеЧисла(3, "вигинтиллиона"),
                    new СклонениеЧисла(4, "вигинтиллиона")
                },
                "вигинтиллионов"),
            new(Степень10.Унвигинт,
                new[] {
                    new СклонениеЧисла(0, "унвигинтиллионов"), new СклонениеЧисла(1, "унвигинтиллион"),
                    new СклонениеЧисла(2, "унвигинтиллиона"), new СклонениеЧисла(3, "унвигинтиллиона"),
                    new СклонениеЧисла(4, "унвигинтиллиона")
                },
                "унвигинтиллионов"),
            new(Степень10.Дуовигинт,
                new[] {
                    new СклонениеЧисла(0, "дуовигинтиллионов"), new СклонениеЧисла(1, "дуовигинтиллион"),
                    new СклонениеЧисла(2, "дуовигинтиллиона"), new СклонениеЧисла(3, "дуовигинтиллиона"),
                    new СклонениеЧисла(4, "дуовигинтиллиона")
                },
                "дуовигинтиллионов"),
            new(Степень10.Тревигинт,
                new[] {
                    new СклонениеЧисла(0, "тревигинтиллионов"), new СклонениеЧисла(1, "тревигинтиллион"),
                    new СклонениеЧисла(2, "тревигинтиллиона"), new СклонениеЧисла(3, "тревигинтиллиона"),
                    new СклонениеЧисла(4, "тревигинтиллиона")
                },
                "тревигинтиллионов"),
            new(Степень10.Кваттуорвигинт,
                new[] {
                    new СклонениеЧисла(0, "кваттуорвигинтиллионов"), new СклонениеЧисла(1, "кваттуорвигинтиллион"),
                    new СклонениеЧисла(2, "кваттуорвигинтиллиона"), new СклонениеЧисла(3, "кваттуорвигинтиллиона"),
                    new СклонениеЧисла(4, "кваттуорвигинтиллиона")
                },
                "кваттуорвигинтиллионов"),
            new(Степень10.Квинвигинт,
                new[] {
                    new СклонениеЧисла(0, "квинвигинтиллионов"), new СклонениеЧисла(1, "квинвигинтиллион"),
                    new СклонениеЧисла(2, "квинвигинтиллиона"), new СклонениеЧисла(3, "квинвигинтиллиона"),
                    new СклонениеЧисла(4, "квинвигинтиллиона")
                },
                "квинвигинтиллионов"),
            new(Степень10.Сексвигинт,
                new[] {
                    new СклонениеЧисла(0, "сексвигинтиллионов"), new СклонениеЧисла(1, "сексвигинтиллион"),
                    new СклонениеЧисла(2, "сексвигинтиллиона"), new СклонениеЧисла(3, "сексвигинтиллиона"),
                    new СклонениеЧисла(4, "сексвигинтиллиона")
                },
                "сексвигинтиллионов"),
            new(Степень10.Септенвигинт,
                new[] {
                    new СклонениеЧисла(0, "септенвигинтиллионов"), new СклонениеЧисла(1, "септенвигинтиллион"),
                    new СклонениеЧисла(2, "септенвигинтиллиона"), new СклонениеЧисла(3, "септенвигинтиллиона"),
                    new СклонениеЧисла(4, "септенвигинтиллиона")
                },
                "септенвигинтиллионов"),
            new(Степень10.Октовигинт,
                new[] {
                    new СклонениеЧисла(0, "октовигинтиллионов"), new СклонениеЧисла(1, "октовигинтиллион"),
                    new СклонениеЧисла(2, "октовигинтиллиона"), new СклонениеЧисла(3, "октовигинтиллиона"),
                    new СклонениеЧисла(4, "октовигинтиллиона")
                },
                "октовигинтиллионов"),
            new(Степень10.Новемвигинт,
                new[] {
                    new СклонениеЧисла(0, "новемвигинтиллионов"), new СклонениеЧисла(1, "новемвигинтиллион"),
                    new СклонениеЧисла(2, "новемвигинтиллиона"), new СклонениеЧисла(3, "новемвигинтиллиона"),
                    new СклонениеЧисла(4, "новемвигинтиллиона")
                },
                "новемвигинтиллионов"),
            new(Степень10.Тригинт,
                new[] {
                    new СклонениеЧисла(0, "тригинтиллионов"), new СклонениеЧисла(1, "тригинтиллион"),
                    new СклонениеЧисла(2, "тригинтиллиона"), new СклонениеЧисла(3, "тригинтиллиона"),
                    new СклонениеЧисла(4, "тригинтиллиона")
                },
                "тригинтиллионов"),
            new(Степень10.Унтригинт,
                new[] {
                    new СклонениеЧисла(0, "унтригинтиллионов"), new СклонениеЧисла(1, "унтригинтиллион"),
                    new СклонениеЧисла(2, "унтригинтиллиона"), new СклонениеЧисла(3, "унтригинтиллиона"),
                    new СклонениеЧисла(4, "унтригинтиллиона")
                },
                "унтригинтиллионов"),
            new(Степень10.Дуотригинт,
                new[] {
                    new СклонениеЧисла(0, "дуотригинтиллионов"), new СклонениеЧисла(1, "дуотригинтиллион"),
                    new СклонениеЧисла(2, "дуотригинтиллиона"), new СклонениеЧисла(3, "дуотригинтиллиона"),
                    new СклонениеЧисла(4, "дуотригинтиллиона")
                },
                "дуотригинтиллионов"),
            new(Степень10.Третригинт,
                new[] {
                    new СклонениеЧисла(0, "третригинтиллионов"), new СклонениеЧисла(1, "третригинтиллион"),
                    new СклонениеЧисла(2, "третригинтиллиона"), new СклонениеЧисла(3, "третригинтиллиона"),
                    new СклонениеЧисла(4, "третригинтиллиона")
                },
                "третригинтиллионов"),
            new(Степень10.Кваттуортригинт,
                new[] {
                    new СклонениеЧисла(0, "кваттуортригинтиллионов"), new СклонениеЧисла(1, "кваттуортригинтиллион"),
                    new СклонениеЧисла(2, "кваттуортригинтиллиона"), new СклонениеЧисла(3, "кваттуортригинтиллиона"),
                    new СклонениеЧисла(4, "кваттуортригинтиллиона")
                },
                "кваттуортригинтиллионов"),
            new(Степень10.Квинтригинт,
                new[] {
                    new СклонениеЧисла(0, "квинтригинтиллионов"), new СклонениеЧисла(1, "квинтригинтиллион"),
                    new СклонениеЧисла(2, "квинтригинтиллиона"), new СклонениеЧисла(3, "квинтригинтиллиона"),
                    new СклонениеЧисла(4, "квинтригинтиллиона")
                },
                "квинтригинтиллионов"),
            new(Степень10.Секстригинт,
                new[] {
                    new СклонениеЧисла(0, "секстригинтиллионов"), new СклонениеЧисла(1, "секстригинтиллион"),
                    new СклонениеЧисла(2, "секстригинтиллиона"), new СклонениеЧисла(3, "секстригинтиллиона"),
                    new СклонениеЧисла(4, "секстригинтиллиона")
                },
                "секстригинтиллионов"),
            new(Степень10.Септентригинт,
                new[] {
                    new СклонениеЧисла(0, "септентригинтиллионов"), new СклонениеЧисла(1, "септентригинтиллион"),
                    new СклонениеЧисла(2, "септентригинтиллиона"), new СклонениеЧисла(3, "септентригинтиллиона"),
                    new СклонениеЧисла(4, "септентригинтиллиона")
                },
                "септентригинтиллионов"),
            new(Степень10.Октотригинт,
                new[] {
                    new СклонениеЧисла(0, "октотригинтиллионов"), new СклонениеЧисла(1, "октотригинтиллион"),
                    new СклонениеЧисла(2, "октотригинтиллиона"), new СклонениеЧисла(3, "октотригинтиллиона"),
                    new СклонениеЧисла(4, "октотригинтиллиона")
                },
                "октотригинтиллионов"),
            new(Степень10.Новемтригинт,
                new[] {
                    new СклонениеЧисла(0, "новемтригинтиллионов"), new СклонениеЧисла(1, "новемтригинтиллион"),
                    new СклонениеЧисла(2, "новемтригинтиллиона"), new СклонениеЧисла(3, "новемтригинтиллиона"),
                    new СклонениеЧисла(4, "новемтригинтиллиона")
                },
                "новемтригинтиллионов"),
            new(Степень10.Квадрагинт,
                new[] {
                    new СклонениеЧисла(0, "квадрагинтиллионов"), new СклонениеЧисла(1, "квадрагинтиллион"),
                    new СклонениеЧисла(2, "квадрагинтиллиона"), new СклонениеЧисла(3, "квадрагинтиллиона"),
                    new СклонениеЧисла(4, "квадрагинтиллиона")
                },
                "квадрагинтиллионов")
        };

        protected static readonly РазрядЧислаНазванияЦифр ЕдиницыРодМуж = new(Разряд.Единицы,
                                                                              new[] {
                                                                                  "ноль", "один", "два", "три",
                                                                                  "четыре", "пять", "шесть", "семь",
                                                                                  "восемь", "девять"
                                                                              });

        protected static readonly РазрядЧислаНазванияЦифр ЕдиницыРодЖен = new(Разряд.Единицы,
                                                                              ЕдиницыРодМуж.Названия
                                                                                           .Select((x, i) => i switch {
                                                                                                1 => "одна",
                                                                                                2 => "две",
                                                                                                _ => x
                                                                                            }).ToArray());

        protected static readonly РазрядЧислаНазванияЦифр Десятки = new(Разряд.Десятки,
                                                                        new[] {
                                                                            "", "десять", "двадцать", "тридцать",
                                                                            "сорок", "пятьдесят", "шестьдесят",
                                                                            "семьдесят", "восемьдесят", "девяносто"
                                                                        });

        protected static readonly РазрядЧислаНазванияЦифр Сотни = new(Разряд.Сотни,
                                                                      new[] {
                                                                          "", "сто", "двести", "триста", "четыреста",
                                                                          "пятьсот", "шестьсот", "семьсот", "восемьсот",
                                                                          "девятьсот"
                                                                      });

        protected static readonly ВалютаСклонения Рубль =
            new(new[] { new СклонениеЧисла(0, "рублей"), new СклонениеЧисла(1, "рубль"), new СклонениеЧисла(2, "рубля"), new СклонениеЧисла(3, "рубля"), new СклонениеЧисла(4, "рубля") },
                "рублей");

        protected static readonly ВалютаСклонения Копейка =
            new(new[] { new СклонениеЧисла(0, "копеек"), new СклонениеЧисла(1, "копейка"), new СклонениеЧисла(2, "копейки"), new СклонениеЧисла(3, "копейки"), new СклонениеЧисла(4, "копейки") },
                "копеек");


        protected static readonly IDictionary<decimal, string> ИменныеЧисла =
            new Dictionary<decimal, string> {
                [0]  = "ноль",
                [10] = "десять",
                [11] = "одиннадцать",
                [12] = "двенадцать",
                [13] = "тринадцать",
                [14] = "четырнадцать",
                [15] = "пятнадцать",
                [16] = "шестнадцать",
                [17] = "семнадцать",
                [18] = "восемнадцать",
                [19] = "девятнадцать"
            };

        protected static IDictionary<decimal, string> АвтоЧислаМуж;
        protected static IDictionary<decimal, string> АвтоЧислаЖен;

        static СправочникЧиселРус() {
            АвтоЧислаМуж = ИменныеЧисла.UnionBy(ЕдиницыРодМуж.Названия
                                                             .Select((ч, i) => (Key: (decimal)i * (byte)Разряд.Единицы,
                                                                                Value: ч))
                                                             .ToDictionary(kv => kv.Key, kv => kv.Value),
                                                kv => kv.Key).UnionBy(Десятки.Названия
                                                                             .Select((ч, i) =>
                                                                                         (Key: (decimal)i * (byte)Разряд.Десятки,
                                                                                          Value: ч))
                                                                             .ToDictionary(kv => kv.Key,
                                                                                           kv => kv.Value),
                                                                      kv => kv.Key).UnionBy(Сотни.Названия
                                                                                                 .Select((ч, i) =>
                                                                                                             (Key: (decimal)i * (byte)Разряд.Сотни,
                                                                                                              Value: ч))
                                                                                                 .ToDictionary(kv => kv.Key,
                                                                                                               kv => kv
                                                                                                                  .Value),
                                                                                            kv => kv.Key)
                                       .ToImmutableSortedDictionary();

            АвтоЧислаЖен = ЕдиницыРодЖен.Названия
                                        .Select((ч, i) => (Key: (decimal)i * (byte)Разряд.Единицы, Value: ч))
                                        .ToDictionary(kv => kv.Key, kv => kv.Value)
                                        .UnionBy(АвтоЧислаМуж, kv => kv.Key)
                                        .ToImmutableSortedDictionary();
        }

        public IDictionary<ВалютаIso, byte> ДробнаяЧастьЧислоЗнаков { get; set; } =
            new Dictionary<ВалютаIso, byte> { [ВалютаIso.Rub] = 2 };

        public string? ДайНазваниеИменногоЧисла(decimal число, bool родМужскойЛи = true) {
            var числа = родМужскойЛи ? АвтоЧислаМуж : АвтоЧислаЖен;

            return числа.TryGetValue(число, out var рез) ? рез : null;
        }

        public string ДайНазваниеКласса(Степень10 степень10, ЧислоЦелоеКласс целоеКласс) {
            if (целоеКласс.Цифры == 0) return "";

            var степень = (Степень10)((byte)степень10 / 3 * 3);
            if (степень > КлассЧислаНазваниеСклонение[^1].Степень10)
                throw new ArgumentException("Столь большое число не поддерживается", nameof(степень10));

            var склСтепени = КлассЧислаНазваниеСклонение.Single(к => к.Степень10 == степень);

            var последРазрядДляСклон = ДайПоследнийРазрядДляСклонения(целоеКласс);

            return склСтепени
                  .Склонения
                  .ElementAtOrDefault(последРазрядДляСклон)
                 ?.Склонение
                ?? склСтепени.Стандартное;
        }


        public bool КлассТребуетЖенРодЛи(Степень10 степень10, ЧислоЦелоеКласс целоеКласс) {
            return степень10 == Степень10.Тыс
                && ЕдиницыРодЖен
                  .Названия
                  .Any(н => н == new ЧислоЦелоеКласс(целоеКласс.РазрядЕдиниц, this).ПреобразуйВТекст(false))
                && целоеКласс.РазрядДесЕд is <= 10 or >= 13;
        }

        public string ДайНазваниеЦифрыРазряда(Разряд разряд, byte цифра, bool родМужскойЛи = true) {
            if (цифра > 9)
                throw new ArgumentException($"Ожидается, что цифра разряда числа будет до 9, но передано {цифра}",
                                            nameof(цифра));
            return разряд switch {
                Разряд.Единицы => родМужскойЛи ? ЕдиницыРодМуж.Названия[цифра] : ЕдиницыРодЖен.Названия[цифра],
                Разряд.Десятки => Десятки.Названия[цифра],
                Разряд.Сотни   => Сотни.Названия[цифра],
                _ => throw new ArgumentException($"Разряд может быть: 1, 10, 100, но передано {разряд}",
                                                 nameof(разряд))
            };
        }

        public string ДайНазваниеБазовойВалютыПолное(ЧислоЦелоеКлассы число,
                                                     ВалютаIso валюта = ВалютаIso.Rub) {
            return валюта switch {
                ВалютаIso.Rub =>
                    Рубль
                       .Склонения
                       .ElementAtOrDefault(ДайПоследнийРазрядДляСклонения(число[Степень10.Едн]))
                      ?.Склонение
                 ?? Рубль.Стандартное,
                _ => throw new ArgumentException($"Валюта с кодом {валюта} не поддерживается", nameof(валюта))
            };
        }

        public string ДайНазваниеБазовойВалютыСокр(ЧислоЦелоеКлассы? число = null,
                                                   ВалютаIso валюта = ВалютаIso.Rub) {
            return валюта switch {
                ВалютаIso.Rub => "руб.",
                _ => throw new ArgumentException($"Валюта с кодом {валюта} не поддерживается", nameof(валюта))
            };
        }

        public string ДайНазваниеДробнойВалютыПолное(ЧислоЦелоеКлассы число,
                                                     ВалютаIso валюта = ВалютаIso.Rub) {
            return валюта switch {
                ВалютаIso.Rub =>
                    Копейка
                       .Склонения
                       .ElementAtOrDefault(ДайПоследнийРазрядДляСклонения(число[Степень10.Едн]))
                      ?.Склонение
                 ?? Копейка.Стандартное,
                _ => throw new ArgumentException($"Валюта с кодом {валюта} не поддерживается", nameof(валюта))
            };
        }

        public string ДайНазваниеДробнойВалютыСокр(ЧислоЦелоеКлассы? число = null,
                                                   ВалютаIso валюта = ВалютаIso.Rub) {
            return валюта switch {
                ВалютаIso.Rub => "коп.",
                _ => throw new ArgumentException($"Валюта с кодом {валюта} не поддерживается", nameof(валюта))
            };
        }

        /// <summary>
        ///     По каким цифрам можно найти склонение числа в справочниках
        /// </summary>
        /// <param name="классЧисла"></param>
        /// <remarks>В русском языке числа первого десятка не требуется склонять</remarks>
        /// <returns>Цифры по которым теоретически можно найти склонение числа или само это число</returns>
        protected static ushort ДайПоследнийРазрядДляСклонения(ЧислоЦелоеКласс классЧисла) {
            return классЧисла.РазрядДесЕд is < 10 or > 20
                       ? классЧисла.РазрядЕдиниц
                       : классЧисла.Цифры;
        }
    }
}