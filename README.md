# Excel Sheet Matcher

Windows desktop utility / десктоп-утилита для Windows — match and find unique rows across Excel, CSV, and Word tables by key columns.

- [English](#english)
- [Русский](#русский)
- [Stack](#stack) · [Requirements](#requirements) · [Build & run](#build--run) · [Project layout](#project-layout) · [License](#license)

---

## English

A Windows desktop utility that builds a combined picture from scattered tabular files: finds **matching** and **unique** rows by user-selected key columns.

The original need was practical — people worked with several Excel/CSV/Word tables and looked for overlapping records by hand. This app loads the files, normalizes keys (including dates), and compares rows using a **precomputed hash** so multi-file matching stays fast.

### Features

- Open multiple files at once: `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.docx`
- Choose key columns (optional RU-style date formats, e.g. `ДД.ММ.ГГГГ`)
- Find **matches** — rows that appear in more than one selected file
- Find **uniques** — rows that appear exactly once
- Replace characters in key fields (option to keep the original)
- Preview a file (double-click in the list)
- Export results to Excel; merge matched rows
- Parse Word documents with splitters (including highlighted text only)



### How it works

1. Files are parsed into rows (`Man` — a header → value dictionary).
2. The user selects key columns; settings live in `KeyHeaderStore`.
3. Each row runs `CalculateHashCode()`: only keys go into the hash (values uppercased; dates normalized).
4. Match/unique search uses `Dictionary<Man, …>` — average lookup ≈ O(1) after hashes are precomputed.
5. Results open in a separate window and can be exported to Excel.



### Usage

1. On the main screen, open files (folder icon).
2. Select the files you need in the list.
3. **Choose key fields** — columns that define when two rows are “the same” record.
4. **Find matches** or **Find uniques**.
5. In the results window, review the grid, optionally merge rows, and export to Excel.

Extra: **Replace characters** — bulk-edit key values before comparison.

### Known limitations

- Only the **first** worksheet of an Excel workbook is read
- Word export is marked as work in progress
- `.xls` and Excel export need Excel installed
- CSV: comma delimiter; quotes are not handled in the current implementation

---



## Русский

Десктоп-утилита для Windows, которая собирает совокупную картину из разрозненных табличных файлов: находит **совпадающие** и **уникальные** строки по выбранным ключевым полям.

Исходная задача была практической — люди работали с несколькими Excel/CSV/Word-таблицами и вручную искали пересечения записей. Программа загружает файлы, нормализует ключи (в том числе даты) и сравнивает строки через **предвычисленный хеш**, чтобы поиск по нескольким файлам оставался быстрым.

### Возможности

- Загрузка нескольких файлов сразу: `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.docx`
- Выбор ключевых колонок (опционально форматы дат в RU-виде, например `ДД.ММ.ГГГГ`)
- Поиск **совпадений** — строки, которые встречаются более чем в одном выбранном файле
- Поиск **уникальных** — строки, которые встречаются ровно один раз
- Замена символов в ключевых полях (с опцией оставить оригинал)
- Просмотр содержимого файла (двойной клик в списке)
- Экспорт результатов в Excel; объединение совпавших строк
- Разбор Word-документов по разделителям (в т.ч. только подсвеченный текст)



### Как это работает

1. Файлы парсятся в набор строк (`Man` — словарь «заголовок → значение»).
2. Пользователь выбирает ключевые поля; настройки хранятся в `KeyHeaderStore`.
3. Для каждой строки вызывается `CalculateHashCode()`: в хеш входят только ключи (значения приводятся к верхнему регистру, даты — к единому формату).
4. Поиск совпадений/уникальных идёт через `Dictionary<Man, …>`: среднее время lookup ≈ O(1) после предрасчёта хешей.
5. Результаты показываются в отдельном окне и при необходимости выгружаются в Excel.



### Использование

1. На главном экране откройте файлы (иконка папки).
2. Отметьте нужные файлы в списке.
3. **Выбрать ключевые поля** — колонки, по которым строки считаются «одной и той же записью».
4. **Найти совпадения** или **Найти уникальные**.
5. В окне результата просмотрите таблицу, при необходимости объедините строки и экспортируйте в Excel.

Дополнительно: **Заменить символы** — массовая правка значений в ключах перед сравнением.

### Известные ограничения

- Читается только **первый** лист книги Excel
- Экспорт в Word помечен как «В разработке»
- Для `.xls` и Excel-экспорта нужен установленный Excel
- CSV: разделитель `,`; кавычки в текущей реализации не обрабатываются

---



## Stack


|                        |                                                          |
| ---------------------- | -------------------------------------------------------- |
| Language               | C#                                                       |
| UI                     | WPF                                                      |
| Runtime                | .NET Framework 4.7.2                                     |
| Excel (modern formats) | ClosedXML                                                |
| Excel `.xls` / export  | Microsoft Excel Interop (COM)                            |
| Word                   | DocumentFormat.OpenXml (+ Interop for unfinished export) |
| CSV                    | `TextFieldParser`                                        |




## Requirements

- Windows
- .NET Framework 4.7.2+
- Visual Studio 2019+ (or MSBuild for .NET Framework) to build
- **Microsoft Excel** installed — required for `.xls` reading and Interop-based export



## Build & run

1. Open `ExcelSheetMatcher.sln` in Visual Studio.
2. Restore NuGet packages (`packages.config`).
3. Build Debug or Release.
4. Run `ExcelSheetMatcher\bin\Debug\ExcelSheetMatcher.exe` (or the Release build).

There is no installer — the built `.exe` and its dependencies are enough.

## Project layout

```
ExcelSheetMatcher/
├── ExcelSheetMatcher.sln
└── ExcelSheetMatcher/
    ├── Views/          # WPF windows (main, keys, results, Word, replace)
    ├── ViewModels/     # Commands and workflows
    ├── Models/         # Man, CellValue, keys, WorksheetItem
    └── Utils/          # Excel/CSV/Word I/O, dialogs
```

Key types:

- `MainViewModel` — file loading, match/unique search
- `Man.CalculateHashCode` — precomputed row hash from keys
- `KeyHeaderStore` — selected key columns
- `WorksheetReader` / `OldWorksheetReader` / `CSVReader` — input parsers
- `WorksheetWriter` — Excel export



## License

Personal project / no license stated. If you publish on GitHub, add an explicit `LICENSE`.