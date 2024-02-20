using System.Reactive.Concurrency;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using ExcelDna.Integration;
using gmafffff.excel.udf.Excel.УправлениеКонтекстом;
using gmafffff.excel.udf.Reactive;

namespace gmafffff.excel.udf.Excel.Команды;

public class ОчередьКомандRx : IDisposable, ИОчередьКоманд
{
    private readonly Dictionary<Type, Func<IEnumerable<ИExcelКоманда>, IEnumerable<ИExcelКоманда>>> _компоновщики =
        new();
    private readonly Subject<ИExcelКоманда> _очередьПодготовки = new();
    private readonly IDisposable _подпискаОчередьИсполнения;

    public readonly IObservable<ИExcelКоманда> ОчередьИсполнения;
    public readonly TimeSpan _периодСбораКоманд = TimeSpan.FromSeconds(.5);

    public ОчередьКомандRx(Action<IList<ИExcelКоманда>>? командаЗапуска = null,
        IScheduler? планировщик = null)
    {
        var упакованныеКоманды = _очередьПодготовки
            .Synchronize()
            // Попытка сгруппировать команды, которые это допускают
            .Where(к => _компоновщики.ContainsKey(к.GetType()))
            .Quiescent(_периодСбораКоманд / 2, планировщик ?? Scheduler.Default)
            .SelectMany(Упаковать);
        
        var одиночныеКоманды = _очередьПодготовки
            .Synchronize()
            .Where(к => !_компоновщики.ContainsKey(к.GetType()));

        ОчередьИсполнения = упакованныеКоманды.Merge(одиночныеКоманды);
        
        _подпискаОчередьИсполнения = ОчередьИсполнения
            .Quiescent(_периодСбораКоманд, планировщик ?? Scheduler.Default)
            .Subscribe(командаЗапуска ?? КомандаЗапускаПоУмолчанию);
    }

    public void ДобавитьКомпоновщикКоманд(Type тип,
        Func<IEnumerable<ИExcelКоманда>, IEnumerable<ИExcelКоманда>> предобработчик) 
        => _компоновщики[тип] = предобработчик;

    public void ДобавитьКоманду(ИExcelКоманда команда) => _очередьПодготовки.OnNext(команда);

    public void Dispose()
    {
        _очередьПодготовки.OnCompleted();
        _подпискаОчередьИсполнения.Dispose();
        _очередьПодготовки.Dispose();
    }

    public static void КомандаЗапускаПоУмолчанию(IList<ИExcelКоманда> команды)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            // using var арп = new АвтоРасчетПриостановить();
            using var оэп = new ОбновлениеЭкранаПриостановить();
            foreach (var команда in команды) команда.Выполнить();
        });
    }

    private IEnumerable<ИExcelКоманда> Упаковать(IList<ИExcelКоманда> команды)
        => from команда in команды
        group команда by команда.GetType()
        into группа
        from упак in _компоновщики[группа.Key](группа)
        select упак;
}