using System.Reactive.Linq;
using System.Reactive.Subjects;
using ExcelDna.Integration;
using gmafffff.excel.udf.Reactive;

namespace gmafffff.excel.udf.AddIn;

public class ExcelМенеджерФоновыхКоманд {
    public static ISubject<ИExcelКоманда> команды = new Subject<ИExcelКоманда>();

    static ExcelМенеджерФоновыхКоманд() {
        // Обработка видимости строк и столбцов
        команды
            .OfType<ExcelКомандаВидимости>()
            .Where(к => к.показать != к.элем.ВидимЛи())
            .BufferWithThrottle(int.MaxValue, TimeSpan.FromSeconds(.5))
            .Subscribe(ПрименитьВидимость);
    }

    public static void ПрименитьВидимость(IList<ExcelКомандаВидимости> команды) {
        var ком_уник = команды
            .Select((к, н) => (к, н))
            .Reverse()
            .DistinctBy(кн => кн.к.элем)
            .OrderBy(кн => кн.н)
            .Select(кн => кн.к)
            .GroupBy(вид => вид.показать)
            .ToArray();
        ExcelAsyncUtil.QueueAsMacro(() => {
            foreach (var гр in ком_уник)
                if (гр.Key)
                    foreach (var ком in гр)
                        ком.элем.Показать(ком.размер);
                else
                    foreach (var ком in гр)
                        ком.элем.Скрыть();
        });
    }

    public static void ЗапланироватьКоманду(ИExcelКоманда команда) {
        команды.OnNext(команда);
    }
}