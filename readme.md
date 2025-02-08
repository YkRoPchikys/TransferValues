# TransferValues - Макрос для переноса данных между Excel файлами

Этот макрос написан на VBA и предназначен для автоматизации процесса переноса данных между Excel файлами. Он позволяет выбрать файл и лист из указанной папки, затем ищет определенные значения в столбце A и переносит найденные данные в ячейки исходной книги.

## Описание работы макроса

1. **Получение списка файлов Excel**:
   - Макрос ищет все файлы с расширением `.xlsx` в указанной папке.
   - Выводится список файлов для выбора пользователем.

2. **Выбор файла для открытия**:
   - Пользователь выбирает файл из списка, который будет открыт для дальнейшей работы.

3. **Выбор листа в выбранном файле**:
   - После открытия файла пользователю предлагается выбрать лист, с которого будут извлечены данные.

4. **Поиск данных в выбранном листе**:
   - Макрос ищет значения в столбце A на выбранном листе.
   - Для каждого поискового термина из заранее заданного списка ищет соответствующее значение в столбце A и извлекает значение из столбца AI (34-й столбец).

5. **Перенос данных в исходную книгу**:
   - После того как значение найдено, оно переносится в заранее заданные ячейки исходной книги.

6. **Закрытие целевого файла**:
   - После завершения переноса данных целевой файл закрывается без сохранения изменений.

7. **Завершение процесса**:
   - После успешного завершения переноса данных выводится сообщение.

## Как использовать

1. **Настройка пути к папке**:
   - Откройте файл с макросом и замените строку:
     ```vba
     folderPath = "C:\Ваш\Путь\К\Папке\" ' ИЗМЕНИТЬ НА СВОЙ ПУТЬ
     ```
     на путь к папке, где находятся файлы Excel, из которых вы хотите переносить данные.

2. **Запуск макроса**:
   - Перейдите на вкладку "Разработчик" в Excel, выберите "Visual Basic", создайте новый модуль и вставьте код макроса.
   - Запустите макрос, и следуйте инструкциям на экране для выбора файла и листа.

## Примечания

- Макрос ожидает, что данные, которые нужно перенести, находятся в столбце A целевой книги.
- Поиск будет производиться по строкам, и макрос завершит поиск после нахождения первого совпадения для каждого поискового термина.
- Макрос закрывает целевой файл без сохранения изменений после завершения работы.
- В случае, если искомое значение не найдено, будет выведено сообщение об ошибке.

## Пример вывода

Если в целевой книге найдены значения, макрос вставит их в заранее заданные ячейки текущего листа:

- Для поискового термина "One" — значение будет вставлено в ячейку `R10` текущей книги.
- Для термина "Two" — значение будет вставлено в ячейку `R15`, и так далее для всех поисковых терминов.

## Лицензия

Этот проект находится в публичном доступе и распространяется под лицензией [MIT License](LICENSE).
