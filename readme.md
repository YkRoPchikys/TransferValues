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

________________________________________________________________________________________________________________

# TransferValues - Macro for Transferring Data Between Excel Files

This macro is written in VBA and is designed to automate the process of transferring data between Excel files. It allows the user to select a file and sheet from a specified folder, then searches for certain values in column A and transfers the found data into cells of the source workbook.

## How the Macro Works

1. **Retrieve List of Excel Files**:
   - The macro searches for all `.xlsx` files in the specified folder.
   - A list of files is displayed for the user to choose from.

2. **Select File to Open**:
   - The user selects a file from the list, which is then opened for further processing.

3. **Select Sheet in the Chosen File**:
   - After opening the file, the user is prompted to select a sheet from which data will be extracted.

4. **Search Data in the Selected Sheet**:
   - The macro searches for values in column A of the selected sheet.
   - For each search term in a predefined list, it searches for the corresponding value in column A and retrieves the value from column AI (the 34th column).

5. **Transfer Data to Source Workbook**:
   - Once a value is found, it is transferred to predefined cells in the source workbook.

6. **Close the Target Workbook**:
   - After transferring the data, the target workbook is closed without saving any changes.

7. **Process Completion**:
   - A message is displayed to inform the user that the data transfer has been completed successfully.

## How to Use

1. **Set the Folder Path**:
   - Open the file containing the macro and modify the following line:
     ```vba
     folderPath = "C:\Your\Path\To\Folder\" ' CHANGE TO YOUR FOLDER PATH
     ```
     Replace it with the path to the folder containing the Excel files you want to transfer data from.

2. **Run the Macro**:
   - Go to the "Developer" tab in Excel, select "Visual Basic", create a new module, and paste the macro code.
   - Run the macro, and follow the on-screen prompts to select the file and sheet.

## Notes

- The macro expects the data to be transferred from column A in the target workbook.
- The search will stop after the first match is found for each search term.
- The macro closes the target workbook without saving changes after the operation is complete.
- If a search term is not found, an error message will be displayed.

## Example Output

If values are found in the target workbook, the macro will insert them into the predefined cells of the source workbook:

- For the search term "One" — the value will be inserted into cell `R10` of the current workbook.
- For the search term "Two" — the value will be inserted into cell `R15`, and so on for all search terms.

## License

This project is open-source and distributed under the [MIT License](LICENSE).

