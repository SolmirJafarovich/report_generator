import json

from src.report_generator.pptx_exporter import ReportGenerator
from src.report_generator.utils.text_config import TextConfig, TextStyle

# --- Загрузка JSON данных ---
with open("report_data.json", encoding="utf-8") as f:
    report_data = json.load(f)

with open("style_config.json", encoding="utf-8") as f:
    style_data = json.load(f)


# --- Преобразование стиля в TextConfig ---
def parse_text_config(data: dict) -> TextConfig:
    return TextConfig(
        title=TextStyle(**data["title"]),
        body=TextStyle(**data["body"]),
        table_header=TextStyle(**data["table_header"]),
        table_cell=TextStyle(**data["table_cell"]),
    )


text_config = parse_text_config(style_data)

# --- Генерация отчёта ---
report = ReportGenerator(report_data, text_config=text_config)
report.build()
report.save("generated_report.pptx")
print("[✓] Отчёт успешно создан: generated_report.pptx")
