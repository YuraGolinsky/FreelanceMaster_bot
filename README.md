The Telegram bot is designed to manage data in an Excel file and perform various operations with data. Here are the main functions and capabilities of the bot: Main functions of the bot 

Link to telegram bot: https://t.me/FreelanceMaster_bot

1. Select Sheet: - Select a sheet from those available in the Excel file. This is necessary for further work with the data.

2. Edit Data: - set the necessary records in the selected sheet by their ID. To edit data, you must enter the record ID, and then new values ​​​​in the format: `UAH Amount, Completed, Payment Date, Money Received`.

3. Add Data: - Set a new record in the selected sheet. Data must be sent in the following format: `ID, UAH Amount, Completed, Payment Date, Money Received`.

4. Calculate the amount: - Calculates the total value of all records in the "UAH Amount" column of the section sheet and displays the result.

5. Delete Data: - View and delete a record by its ID on a two-sheet sheet.

6. Menu Navigation: - Users can navigate to the main menu or perform various actions via the menu buttons.

Interface and Controls - Main Menu: - `📄 Select Sheet` - To select an Excel sheet.
- `✏️ Edit Data` - To edit existing records.
- `➕ Add Data` - To add new records.
- `💰 Calculate Sum` - To calculate the total amount.
- `🗑️ Delete Data` - To delete records.

- Action Submenu: - `✏️ Edit Data` - Starts the data editing process.
- `➕ Add Data` - Starts the data adding process.
- `💰 Calculate Sum` - Performs the sum calculation.
- `🗑️ Delete Data` - Starts the data deletion process.
- `🏠 Back to Main Menu` - Returns the user to the main menu.

Usage Instructions 1. Running the Bot: - Run the bot by running `python bot.py` after installing all dependencies.

2. Interacting with the Bot: - Open Telegram and navigate to your bot.
- Use the chat commands and buttons to perform various operations.

3. Input Data: - When adding or editing data, make sure the data format is correct and meets the requirements (e.g. date format `DD.MM.YYYY`).

4. Error Handling: - The bot provides error messages if something went wrong, such as invalid data format or file access issues.

Technical Details - Excel File: - The bot works with an Excel file, the path to which is specified in the `FILE_PATH` variable. Make sure the file exists and is readable/writeable.

- Data Format: - The file must save a table with the fields ID, Amount UAH, Completed, Payment date, Money received.

- API Token: - It is necessary to replace TOKEN in the token code provided by BotFather when creating the bot.

5. Preparing an Excel file Make sure that the Excel file with the name is located in the specified path: . This file should save data in the following format: | ID | Amount UAH | Completed | Payment date | Money received |
|------|-----------|-------------|---- -----------|
| 1 | 500 | ✅ | 12.08.2024 | ❌ |
| 2 | 200 | ❌ | 15.08.2024 | ✅ |

6. Usage examples

- Edit data:
- Click "✏️ Edit data", then enter the record ID.
- Enter new data in the format `UAH Amount, Completed, Payment date, Money received`.

- Add data:
- Click "➕ Add data" and send data in the format `ID, UAH Amount, Completed, Payment date, Money received`.

- Delete data:
- Click "🗑️ Delete data" and enter the record ID to delete.

- Calculate the amount:
- Click "💰 Calculate the amount" to get the total amount of all records.

![image](https://github.com/user-attachments/assets/5a5800f7-491e-4249-b718-858de2703a06)

![image](https://github.com/user-attachments/assets/7cce606a-fff5-46ee-bc2a-4220ba8eeec4)

![image](https://github.com/user-attachments/assets/4ec769ce-3fde-4785-818d-015e76d17136)












![image](https://github.com/user-attachments/assets/8f4e1d23-57ca-47d8-834c-06ab7d22c206)

![image](https://github.com/user-attachments/assets/4ba4bcab-8fa0-4723-9785-a9a4c1995310)



![image](https://github.com/user-attachments/assets/5cd93a01-bc8f-4929-970a-8de2883b9cac)





![image](https://github.com/user-attachments/assets/dae71994-f77f-41c2-9477-5b2d0235dc3a)

![image](https://github.com/user-attachments/assets/af58e56b-3a86-4545-90fe-d3f19f083d47)

Translation into Russian


Telegram-бот предназначен для управления данными в Excel - файле и выполнения различных операций с данными. Вот основные функции и возможности бота: Основные функции бота 

Ссылка на телеграм бот: https://t.me/FreelanceMaster_bot

1. Выбор Листа: - Выберите лист из доступного в Excel-файле. Это необходимо для дальнейшей работы с данными.

2. Редактирование Данных: - установить необходимые записи в выбранном листе по их идентификатору. Для редактирования данных необходимо ввести идентификатор записи, а затем новые значения в формате: `Сумма грн, Выполнил, Дата выплаты, Пришли деньги`.

3. Добавление Данных: - Установите новую запись в выбранный лист. Данные необходимо отправить в следующем формате: `ID, Сумма грн, Выполнил, Дата выплаты, Пришли деньги`.

4. Вычисление суммы: - Подсчитывает общую величину всех записей в столбце "Сумма грн" листа секции и выводит результат.

5. Удаление Данных: - Посмотрите и удалите запись по ее удостоверению личности на двухлистовом листе.

6. Навигация по меню: - Пользователи могут переходить в главное меню или переходить к различным действиям через кнопки меню.

 Интерфейс и Управление - Главное Меню: - `📄 Выбрать лист` - Для выбора листа Excel.
 - `✏️ Редактировать данные` - Для редактирования существующих записей.
 - `➕ Добавить данные` - Для добавления новых записей.
 - `💰 Рассчитать сумму` - Для расчета общей суммы.
 - `🗑️ Удалить данные` - Для удаления записей.

- Подменю Действий: - `✏️ Редактировать данные` - Начинает процесс редактирования данных.
 - `➕ Добавить данные` - Начинает процесс добавления данных.
 - `💰 Рассчитать сумму` - Выполняет расчет суммы.
 - `🗑️ Удалить данные` - Начинает процесс удаления данных.
 - `🏠 Назад в главное меню` - Возвращает пользователя в главное меню.

 Инструкции по Использованию 1. Запуск Бота: - Запустите бот, выполнив команду `python bot.py` после установки всех зависимостей.

2. Взаимодействие с Ботом: - Откройте Telegram и перейдите к вашему боту.
 - Используйте команды и кнопки чата для выполнения различных операций.

3. Входные данные: - При добавлении или редактировании данных убедитесь, что формат данных правильный и соответствует требованиям (например, формат даты `ДД.ММ.ГГГГ`).

4. Обработка Ошибок: — Бот предоставляет сообщения об ошибках, если что-то пошло не так, например, неверный формат данных или проблемы с доступом к файлу.

 Технические детали - Файл Excel: - Бот работает с Excel-файлом, путь к которому указан в переменной `FILE_PATH`. Убедитесь, что файл существует и доступен для чтения/записи.

- Формат Данных: - Файл должен сохранять таблицу с полями ID, Сумма грн, Выполнил, Дата выплаты, Пришли деньги.

- Токен API: - Необходимо заменить TOKEN в коде токена, предоставленном BotFather при создании бота.

5. Подготовка Excel-файла Убедитесь, что файл Excel с именем находится по указанному пути: . Этот файл должен сохранять данные в формате: | удостоверение личности | Сумма грн | Выполнил | Дата выплаты | Пришли деньги |
|------|-----------|----------|--------------|---- -----------|
| 1 | 500 | ✅ | 12.08.2024 | ❌ |
| 2 | 200 | ❌ | 15.08.2024 | ✅ |

 6. Примеры Использования

- Редактирование данных:
  - Нажмите "✏️ Редактировать данные", затем введите ID записи.
  - Введите новые данные в формате `Сумма грн, Выполнил, Дата выплаты, Пришли деньги`.

- Добавление данных:
  - Нажмите "➕ Добавить данные" и отправьте данные в формате `ID, Сумма грн, Выполнил, Дата выплаты, Пришли деньги`.

- Удаление данных:
  - Нажмите "🗑️ Удалить данные" и введите ID записи для удаления.

- Расчет суммы:
  - Нажмите "💰 Рассчитать сумму" для получения общей суммы всех записей.
