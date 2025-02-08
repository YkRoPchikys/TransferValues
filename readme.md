# TransferValues - Макрос для переноса данных между Excel файлами

Этот макрос написан на VBA и предназначен для автоматизации процесса переноса данных между Excel файлами. Он позволяет выбрать файл и лист из указанной папки, затем ищет определенные значения в столбце A и переносит найденные данные в ячейки исходной книги.

## Как использовать

### Импорт макроса в Excel

1. **Откройте Excel**.
2. Перейдите на вкладку **Разработчик**. Если у вас эта вкладка не отображается, включите её:
   - Нажмите **Файл** > **Параметры**.
   - В разделе **Настроить ленту** установите флажок **Разработчик**.
3. Нажмите **Visual Basic** на вкладке **Разработчик**.
4. В редакторе VBA выберите **Вставка** > **Модуль** для создания нового модуля.
5. Скопируйте код макроса из этого репозитория и вставьте его в новый модуль.
6. Закройте редактор VBA и вернитесь в Excel.
7. Перейдите на вкладку **Разработчик** и нажмите **Макросы**.
8. Выберите макрос `TransferValues` и нажмите **Запустить**.

### Импорт в Личную книгу

Если вы хотите использовать макрос в любой книге Excel без необходимости повторного импорта, вы можете добавить его в Личную книгу:

1. Откройте Excel.
2. Нажмите **Alt + F11**, чтобы открыть редактор VBA.
3. Перейдите в **Личную книгу (PERSONAL.XLSB)**. Если её нет, создайте её:
   - Нажмите **Файл** > **Новый** и выберите **Личная книга макросов**.
4. В редакторе VBA выберите **Вставка** > **Модуль**.
5. Скопируйте код макроса из этого репозитория и вставьте его в новый модуль.
6. Сохраните и закройте редактор.
7. После этого макрос будет доступен в любой новой книге Excel.

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

## How to Use

### Import the Macro into Excel

1. **Open Excel**.
2. Go to the **Developer** tab. If you don't see this tab, enable it by:
   - Clicking **File** > **Options**.
   - Under **Customize Ribbon**, check the **Developer** box.
3. Click **Visual Basic** on the **Developer** tab.
4. In the VBA editor, select **Insert** > **Module** to create a new module.
5. Copy the macro code from this repository and paste it into the new module.
6. Close the VBA editor and return to Excel.
7. Go to the **Developer** tab and click **Macros**.
8. Select the `TransferValues` macro and click **Run**.

### Import into Personal Workbook

If you want to use the macro in any Excel workbook without re-importing it, you can add it to your Personal Workbook:

1. Open Excel.
2. Press **Alt + F11** to open the VBA editor.
3. Go to **Personal Workbook (PERSONAL.XLSB)**. If it doesn't exist, create it:
   - Click **File** > **New** and select **Personal Macro Workbook**.
4. In the VBA editor, select **Insert** > **Module**.
5. Copy the macro code from this repository and paste it into the new module.
6. Save and close the editor.
7. After this, the macro will be available in any new Excel workbook.

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


