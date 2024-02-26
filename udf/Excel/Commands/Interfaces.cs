namespace gmafffff.excel.udf.Excel.Команды;

public interface ИExcelКоманда {
    void Выполнить();
}

public interface ИОчередьКоманд {
    void ДобавитьКоманду(ИExcelКоманда команда);

    void ДобавитьКомпоновщикКоманд(Type тип,
                                   Func<IEnumerable<ИExcelКоманда>, IEnumerable<ИExcelКоманда>> предобработчик);
}