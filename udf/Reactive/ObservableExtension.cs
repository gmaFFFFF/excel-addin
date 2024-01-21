using System.Reactive;
using System.Reactive.Linq;

namespace gmafffff.excel.udf.Reactive;

public static class ObservableExtension {
    // https://stackoverflow.com/questions/35557411/buffer-until-quiet-behavior-from-reactive
    public static IObservable<IList<TSource>> BufferWithThrottle<TSource>(this IObservable<TSource> source,
        int maxAmount, TimeSpan порог) {
        return Observable.Create<IList<TSource>>(obs => {
            return source.GroupByUntil(_ => true,
                    g => g.Throttle(порог).Select(_ => Unit.Default)
                        .Merge(g.Take(maxAmount)
                            .LastAsync()
                            .Select(_ => Unit.Default)))
                .SelectMany(i => i.ToList())
                .Subscribe(obs);
        });
    }
}