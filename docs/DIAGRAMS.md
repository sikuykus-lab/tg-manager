# Диаграммы

Три вида схемы — как в [dataroom-cms](https://github.com/sikuykus-lab/dataroom-cms):
**данные**, **взаимодействие пользователя**, **процессы администратора**.

Рендер: скопировать блок в [mermaid.live](https://mermaid.live).

## Схема данных

```mermaid
flowchart TB
  subgraph source ["Источник"]
    SH["Лист Экспорт\nGoogle Sheets"]
  end

  subgraph cache ["Кэш бота"]
    DB["SQLite data.db"]
  end

  subgraph ui ["Telegram"]
    SR["Поиск по блоку"]
    CD["Карточка вехи"]
  end

  SH -->|"paste / sync"| DB
  DB --> SR
  DB --> CD
```

## Процесс пользователя

```mermaid
flowchart LR
  A["Нужна веха"] --> B{"Знаю блок?"}
  B -->|да| C["Список вех"]
  B -->|блок+шифр| D["Карточка +\nзависимости"]
  C --> E["Клик → детали"]
  D --> F["Решение на площадке"]
```

## Процессы администратора

```mermaid
flowchart TD
  R1["Свежий Экспорт\nиз таблицы"] --> R2["Ctrl+C → боту"]
  R2 --> R3["data.db обновлён"]
  R3 --> R4["Поиск сразу\nновые даты"]
```
