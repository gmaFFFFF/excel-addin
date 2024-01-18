# Add-in (надстройка) для Excel (в разработке)

Add-in построен на базе [Excel-DNA](https://excel-dna.net).

## Добавляет пользовательские функции

| Название           | Описание                                                                                                                                     |
|--------------------|----------------------------------------------------------------------------------------------------------------------------------------------|
|                    | **Числовые функции**:                                                                                                                        |
| РублиПрописью      | Отображает сумму в рублях прописью                                                                                                           |
| ОкруглГаус         | Округление десятичных чисел по Гауссу (до ближайшего четного знака)                                                                          |
|                    | **JSON**:                                                                                                                                    |
| JsonИндекс         | Извлекает элементы json по индексу                                                                                                           |
| JsonPath           | Извлекает элементы json с помощью синтаксиса[JSONPath](https://danielaparker.github.io/JsonCons.Net/articles/JsonPath/JsonConsJsonPath.html) |
| JmesPath           | Извлекает элементы json с помощью синтаксиса[JMESPath](https://jmespath.org/specification.html)                                              |
|                    | **Http**:                                                                                                                                    |
| HttpGet_active     | Get запрос (выполняется каждый раз при пересчете формулы :confused:)                                                                         |
| HttpPost_active    | Post запрос (выполняется каждый раз при пересчете формулы :confused:)                                                                        |
| Base64Кодировать   | Кодирует текст в base64 код                                                                                                                  |
| Base64Декодировать | Декодирует текст из формата base64                                                                                                           |

## Установка надстройки

### Установка для пользователя системы

Согласно [инструкции](https://support.microsoft.com/ru-ru/office/добавление-и-удаление-надстроек-в-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) от Microsoft.

### Загрузка надстройки при запуске только конкретного файла Excel

Сохраните файл как книгу Excel с поддержкой макросов (*.xlsm).
Разместите рядом с файлом Excel файлы надстройки "*.xll"
В проекте VBA добавьте к элементу `ЭтаКнига` следующую процедуру:

```
Private Sub Workbook_Open()
Dim succes As Boolean
Dim addinPath, addinName, addinFullName, defaultPath As String

    addinPath = IIf(Not ActiveWorkbook Is Nothing, ActiveWorkbook.Path, _
                IIf(Not ActiveWindow.Parent Is Nothing, ActiveWindow.Parent.Path, _
                    ThisWorkbook.Path))
      
    addinName = "gmafffff.excel.udf"
    #If Win64 Then
        Debug.Print ("x64")
        addinName = addinName & "_x64.xll"
    #Else
        Debug.Print ("x32")
        addinName = addinName & "_x32.xll"
    #End If
    addinFullName = addinPath & "\" & addinName
   
    defaultPath = Application.DefaultFilePath
    Application.DefaultFilePath = addinPath

    For Each надстр In Application.AddIns2
        If надстр.Name = addinName Then
            Exit Sub
        End If
    Next

    Debug.Print ("Load addin")
    succes = Application.RegisterXLL(addinFullName)

    If succes Then
        Debug.Print ("addin " & addinName & " успешно загружен")
    Else
        Debug.Print ("addin " & addinName & " не загружен")
    End If

    Application.DefaultFilePath = defaultPath
End Sub
```