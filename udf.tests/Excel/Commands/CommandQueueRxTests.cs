using AutoFixture;
using AutoFixture.AutoMoq;
using FluentAssertions;
using FluentAssertions.Execution;
using gmafffff.excel.udf.Excel.Команды;
using Microsoft.Reactive.Testing;

namespace udf.tests.Excel.Команды;

public class ОчередьКомандRxТесты {
    private readonly Fixture fixture = new();

    public ОчередьКомандRxТесты() {
        fixture.Customize(new AutoMoqCustomization { ConfigureMembers = true });
        fixture.Behaviors.OfType<ThrowingRecursionBehavior>().ToList()
               .ForEach(b => fixture.Behaviors.Remove(b));
        fixture.Behaviors.Add(new OmitOnRecursionBehavior());
        fixture.Inject(new IntPtr(0));
    }

    [Fact]
    public void ПланировщикПередаетНаИсполнениеПоступившиеКоманды() {
        // Подготовка
        var             команды     = fixture.CreateMany<ИзмениВидимостьРядаКоманда>();
        var             планировщик = new TestScheduler();
        ОчередьКомандRx очередь     = new(_ => { }, планировщик);

        var observerВыход = планировщик.CreateObserver<ИExcelКоманда>();
        очередь.ОчередьИсполнения.Subscribe(observerВыход);

        // Действие
        foreach (var команда in команды) очередь.ДобавитьКоманду(команда);
        планировщик.Start();

        // Проверка
        using var пространствоПроверки = new AssertionScope();
        var полученыКоманды = from m in observerВыход.Messages
                              select m.Value.Value;
        полученыКоманды.Should().BeEquivalentTo(команды, "планировщик должен вернуть то, что поставлено в очередь");
        планировщик.Clock.Should()
                   .BeGreaterOrEqualTo(очередь._периодСбораКоманд.Ticks, "должно пройти время на сборку команд");
    }

    [Fact]
    public void ПланировщикУмеетГруппироватьКоманды() {
        // Подготовка
        var             команды      = fixture.CreateMany<ИзмениВидимостьРядаКоманда>();
        var             группаКоманд = ИзмениВидимостьРядаКоманда.Упаковать(команды);
        var             планировщик  = new TestScheduler();
        ОчередьКомандRx очередь      = new(_ => { }, планировщик);
        очередь.ДобавитьКомпоновщикКоманд(typeof(ИзмениВидимостьРядаКоманда), ИзмениВидимостьРядаКоманда.Упаковать);

        var observerВыход = планировщик.CreateObserver<ИExcelКоманда>();
        очередь.ОчередьИсполнения.Subscribe(observerВыход);

        // Действие
        foreach (var команда in команды) очередь.ДобавитьКоманду(команда);
        планировщик.Start();

        // Проверка
        using var пространствоПроверки = new AssertionScope();
        var полученыКоманды = from m in observerВыход.Messages
                              select m.Value.Value;
        полученыКоманды.Should().BeEquivalentTo(группаКоманд, "планировщик должен вернуть сгруппированные команды");
        планировщик.Clock.Should()
                   .BeGreaterOrEqualTo(очередь._периодСбораКоманд.Ticks, "должно пройти время на сборку команд");
    }
}