using System.Reactive;
using System.Reactive.Concurrency;
using System.Reactive.Linq;

namespace gmafffff.excel.udf.Reactive;

public static class ObservableExtension {
    // https://stackoverflow.com/questions/35557411/buffer-until-quiet-behavior-from-reactive
    /// <summary>
    ///     Аналог для <see cref="Quiescent{T}" />
    /// </summary>
    /// <param name="source"></param>
    /// <param name="порог"></param>
    /// <param name="maxAmount"></param>
    /// <typeparam name="TSource"></typeparam>
    /// <returns></returns>
    public static IObservable<IList<TSource>> BufferWithThrottle<TSource>(this IObservable<TSource> source,
                                                                          TimeSpan порог,
                                                                          int maxAmount = int.MaxValue) {
        return Observable.Create<IList<TSource>>(obs
                                                     => source.GroupByUntil(_ => true,
                                                                            g => g.Throttle(порог)
                                                                                  .Select(_ => Unit.Default)
                                                                                  .Merge(g.Take(maxAmount)
                                                                                          .LastAsync()
                                                                                          .Select(_ => Unit.Default)))
                                                              .SelectMany(i => i.ToList())
                                                              .Subscribe(obs));
    }

    // LINQ Operators and Composition in https://introtorx.com/chapters/key-types 
    /// <summary>
    ///     Буферизует наблюдаемую последовательность пока после последнего выпущенного элемента не пройдет
    ///     время <see cref="minimumInactivityPeriod" />
    /// </summary>
    /// <param name="src"></param>
    /// <param name="minimumInactivityPeriod">максимальный период между двумя выпусками, которые попадут в один буфер</param>
    /// <param name="scheduler">Планировщик</param>
    /// <typeparam name="T">Тип наблюдаемых элементов</typeparam>
    /// <returns>Буферизованная последовательность</returns>
    public static IObservable<IList<T>> Quiescent<T>(this IObservable<T> src,
                                                     TimeSpan minimumInactivityPeriod, IScheduler scheduler) {
        var onoffs =
            from _ in src
            from delta in
                Observable.Return(1, scheduler)
                          .Concat(Observable.Return(-1, scheduler)
                                            .Delay(minimumInactivityPeriod, scheduler))
            select delta;
        var outstanding   = onoffs.Scan(0, (total, delta) => total + delta);
        var zeroCrossings = outstanding.Where(total => total == 0);
        return src.Buffer(zeroCrossings);
    }
}