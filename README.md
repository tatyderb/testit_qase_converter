# testit_qase_converter
Conversion from xlsx testit exported project into qase tests

1. Export testit project as xlsx file with the one worksheet.
2. Convert xlsx into json
3. Import json into qase.

# Терминология TestIt и Qase

| Qase       | TestIt    | xlsx header  | json path                            | Комментарий                                             |
|------------|-----------|--------------|--------------------------------------|---------------------------------------------------------|
| Repository | Project   | нет          | нет                                  | импортируем и экспортируем руками в текущий репозиторий |
| Suites     | Direction | Direction    | `d['suites'][i]['title']`            | `d['suites'][i]['suites']` описывает вложенные разделы  |
| Case       | TestCase  | TestCaseName | `d['suites'][i]['cases'][j]['title']` | id suites сквозная, id cases - своя, сквозная          |

# TestIt xlsx - столбцы экспортируемого файла

## Строки с заполненным столбцом Id

```python
{
    'Id': '341',     # id теста, не нужно при импорте (может, будут какие-то ссылки потом на него?
    'Direction': 'Demo Project',   # название проекта 
    'Section': 'Регистрация пользователя', # секция в проекте
    'TestCaseName': 'Регистрация с валидными данными через кнопку "Регистрация" в хэдере',  # название проекта
    'Automated': 'False', 
    'Preconditions': None,  # всегда пустая
    'Steps': None,          # всегда пустая
    'Postconditions': None, # всегда пустая
    'ExpectedResult': None, # всегда пустая
    'TestData': None, 
    'Comments': None, 
    'Iterations': None, 
    'Priority': 'Medium', 
    'State': 'NotReady', 
    'CreatedDate': '07/24/2022 10:54:44', 
    'CreatedById': 'UnknownUser', 
    'Tags': None
}
```

## Строки с незаполненным столбцом Id

Эти строки предназначены для описания шагов Preconditions, Steps, Postconditions. 
При этом действия пишутся в указанных столбцах, а ожидаемый результат - всгда в столбце ExpectedResult.

# Qase json формат

Был получен сначала экспортом существующих тестов (из демо проекта), потом импортом этих же тестов в другой проект.

##