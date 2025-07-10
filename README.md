# Report Generator

**Report Generator** — это модульный Python-инструмент для генерации презентаций PowerPoint на основе шаблона и JSON-конфигураций.
Он позволяет создавать структурированные отчёты с текстом, изображениями и таблицами,
используя заранее определённые стили и шаблоны.


---

###  Примеры использования

#### 1. Генерация отчёта из CLI:

```bash
poetry run python report_generate.py --data report_data.json --style style_config.json --output report.pptx
```

* `report_data.json` — описание содержимого слайдов
* `style_config.json` — конфигурация стилей текста
* `report.pptx` — путь к сохраняемому файлу отчёта

> Примеры использования находятся в `examples.ipynb`

---

###  Типы слайдов

Поддерживаются:

| Тип `type` | Описание                                      |
| ---------- | --------------------------------------------- |
| `title`    | Титульный слайд с заголовком и подзаголовком  |
| `text`     | Текстовый слайд (заголовок + контент + notes) |
| `image`    | Слайд с изображениями (до 4х на слайд)        |
| `table`    | Адаптивный слайд с таблицей                   |


---

###  Конфигурация

####  `report_data.json`

```json
{
  "template_path": "templates/example_template.pptx",
  "slides": [
    {
      "type": "image",
      "title": "Пример изображения",
      "images": [
        {"path": "data/image1.jpg", "caption": "Картинка 1"}
      ]
    },
    {
      "type": "table",
      "title": "Результаты",
      "tables": [
        {"path": "data/table1.xlsx"}
      ]
    }
  ]
}
```

####  `style_config.json`

```json
{
  "title": {
    "font_name": "Calibri",
    "font_size": 28,
    "bold": true,
    "italic": false,
    "color": "002060",
    "text_align": "center"
  },
  "body": {
    "font_name": "Calibri",
    "font_size": 16,
    "bold": false,
    "italic": false,
    "color": "333333"
  },
  "table_header": {
    "font_name": "Calibri",
    "font_size": 14,
    "bold": true,
    "color": "000000"
  },
  "table_cell": {
    "font_name": "Calibri",
    "font_size": 12,
    "bold": false,
    "color": "000000"
  }
}
```
