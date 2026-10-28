# RMS DMN Parser

Парсер DMN-файлов в читаемую Excel-книгу для системных аналитиков.

## Что делает

- рекурсивно читает все `.dmn` из корня проекта и его подпапок (включая
  `dmn_for_parser`);
- выгружает decision tables в табличном виде;
- поддерживает таблицы, вложенные в `contextEntry`, и таблицы из
  `businessKnowledgeModel/encapsulatedLogic`;
- показывает DMN-выражения, если в decision нет таблицы;
- раскрывает файлы-словари и файлы типов через `itemDefinition`;
- собирает общие листы `Summary`, `Inputs`, `Types`, `Dependencies`;
- продолжает обработку остальных файлов, если один DMN оказался проблемным.

## Установка

```powershell
py -3.11 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

На macOS/Linux активация окружения выглядит так:

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Если окружение уже создано как `.venv-win`, можно использовать его локально на Windows.

## Использование

Положите `.dmn` файлы рядом со скриптом или в папку:

```text
dmn_for_parser
```

Подпапки также обходятся рекурсивно. Скрытые служебные каталоги (например,
`.git`) пропускаются.

Запустите:

```powershell
python dmn_to_table.py
```

Или через локальное окружение проекта:

```powershell
.venv-win\Scripts\python.exe dmn_to_table.py
```

Результат появится в корне проекта:

```text
result_ГГГГ-ММ-ДД_ЧЧ-ММ-СС.xlsx
```

## Структура

```text
helper_rms/
├── dmn_to_table.py
├── dmn_for_parser/
├── requirements.txt
└── README.md
```

Транскрайбер вынесен в отдельный проект `../rms_transcriber`.
