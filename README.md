# Телефонная книга 📖 (Windows Forms)

Приложение для учета контактов налогоплательщиков с возможностью фильтрации (ФЛ/ЮЛ), оценок (лайки/дизлайки) и экспорта в Excel.

![Скриншот интерфейса](https://via.placeholder.com/800x500?text=Screenshot+of+PhoneBook+App)  
*(Пример изображения — замените на реальный скриншот)*

## 🔧 Технологии
- **C#** (.NET Framework, Windows Forms)
- **SQL Server** (хранение данных)
- **Excel Interop** (экспорт в XLSX)
- **Асинхронные операции** (для плавного UI)

## 📦 Установка
1. Клонируйте репозиторий:
   ```bash
   https://github.com/Resoulone1993/Telephone.git

    Функционал
CRUD-операции (добавление/редактирование).

Фильтрация:

Поиск по всем полям (ИНН, ФИО, телефон и др.).

Отдельные вкладки для Физлиц (ИНН > 11 цифр) и Юрлиц (ИНН < 11 цифр).

Оценки:

Лайки/дизлайки с сохранением в БД.

Экспорт в Excel (автосохранение в D:\Файлы телефонной книги).

📊 Структура БД
Таблица telefon_table:

sql
Copy
CREATE TABLE [telefon_table] (
    [Id] INT IDENTITY(1,1) PRIMARY KEY,
    [INN] NVARCHAR(20),
    [Name] NVARCHAR(100),
    [Tel] NVARCHAR(20),
    [Date] DATETIME,
    [inspecter] NVARCHAR(50),
    [Otdel] NVARCHAR(50),
    [Polzovatel] NVARCHAR(50),
    [inlike] INT DEFAULT 0,
    [Dislike] INT DEFAULT 0
);

📜 Лицензия
MIT License. Подробнее см. в файле LICENSE.
