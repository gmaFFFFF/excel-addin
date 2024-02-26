namespace gmafffff.excel.udf.ФИО;

public class ФИО {
// автор оригинального кода - Aent(Андрей Энтелис)
// http://www.programmersforum.ru/showpost.php?p=757147&postcount=6
// https://excelvba.ru/code/CropFIO
    public static string СократитьФио(string s, bool слева = false) {
        s = s.Trim();
        var sФ = "";
        var sИ = "";
        var sО = "";
        int k;

        var cultureInfo = Thread.CurrentThread.CurrentCulture;
        var textInfo    = cultureInfo.TextInfo;

        // Инициалы заданы явно или пустая строка
        if (s.Length == 0 || s.IndexOf(".") > 0) return s;

        // Нормализация входной строки
        s = s.Replace(((char)30).ToString(), "-")
             .Replace(" -", "-")
             .Replace("- ", "-")
             .Replace("' ", "'")
             .Replace(" '", "'"); // О 'Генри Александр; О' Генри Александр; Н' Гомо; Д' Тревиль

        var sv = s.Split();

        var i = sv.Length - 1;
        if (i < 1) return s;

        switch (sv[i - 1]) {
            case "оглы":
            case "кызы":
            case "заде":
            {
                // бей, бек, заде, зуль, ибн, кызы, оглы, оль, паша, уль, хан, шах, эд, эль
                i--;
                sО = sv[i - 1][..1].ToUpper() + ".";
                i--;
                break;
            }

            case "паша":
            case "хан":
            case "шах":
            case "шейх":
            {
                i--;
                break;
            }

            default:
            {
                switch (sv[i][^3..]) {
                    case "вич":
                    case "вна":
                    case "ной":
                    case "чем":
                    case "ича":
                    case "ичу":
                    case "вны":
                    case "вне":
                    {
                        if (i >= 2) {
                            sО = СropWord(sv[i]);
                        }
                        else {
                            sИ = СropWord(sv[i]);
                            sФ = sv[0];
                        }

                        i -= 1;
                        break;
                    }

                    default:
                    {
                        k = sv[i].IndexOf("-");
                        if (k > 0) {
                            switch (sv[i][k..]) {
                                case "оглы":
                                case "кызы":
                                case "заде":
                                case "угли":
                                case "уулы":
                                case "оол":
                                {
                                    // Вариант насаба «-оглы» и «-заде»  типа Махмуд-оглы
                                    sО =  sv[i][..1].ToUpper() + ".";
                                    i  -= 1;
                                    if (i == 0) {
                                        sИ = sО;
                                        sО = "";
                                    }

                                    break;
                                }
                            }
                        }
                        else if (i > 2) {
                            switch (sv[i - 1]) {
                                case "ибн":
                                case "бен":
                                case "бин":
                                {
                                    sО =  sv[i][..1].ToUpper() + "."; // Усерталь Алишер бен Сулейман
                                    i  -= 2;
                                    break;
                                }
                            }
                        }
                        else {
                            sИ = sv[i][..1].ToUpper();
                            if (sv[i].Length > 1) sИ += ".";
                            i--;
                        }

                        break;
                    }
                }

                break;
            }
        }


        switch (sv[0]) {
            case "де":
            case "дел":
            case "дос":
            case "cент":
            case "ван":
            case "фон":
            case "цу":
            {
                if (i >= 2) {
                    sФ = sv[0] + " " + textInfo.ToTitleCase(sv[1]);
                    sИ = СropWord(sv[2]);
                }
                else if (sИ.Length > 0) {
                    sФ = sv[0] + " " + textInfo.ToTitleCase(sv[1]);
                }
                else {
                    sФ = textInfo.ToTitleCase(sv[1]);
                    sИ = СropWord(sv[1]);
                }

                break;
            }

            default:
            {
                if (sФ.Length == 0) {
                    sФ = textInfo.ToTitleCase(sv[0]);
                    if (sИ.Length == 0) sИ = СropWord(sv[1]);
                }

                break;
            }
        }

        return слева
                   ? $"{sИ}{sО} {sФ}"
                   : $"{sФ} {sИ}{sО}";
    }

    private static string СropWord(string s) {
        if (s.Length == 1) return s;
        var ss        = s[..1].ToUpper() + ".";
        var k         = s.IndexOf("-");
        if (k > 0) ss = $"{ss}-{s[k..1]}.";
        return ss;
    }
}