import argparse
import json
from pathlib import Path

from src.report_generator.pptx_exporter import ReportGenerator
from src.report_generator.utils.text_config import TextConfig, TextStyle


def load_json(path: Path) -> dict:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def parse_text_config(data: dict) -> TextConfig:
    return TextConfig(
        title=TextStyle(**data.get("title", {})),
        body=TextStyle(**data.get("body", {})),
        table_header=TextStyle(**data.get("table_header", {})),
        table_cell=TextStyle(**data.get("table_cell", {})),
        notes=TextStyle(**data.get("notes", {})),
    )


def main():
    parser = argparse.ArgumentParser(
        description="Генерация отчёта PowerPoint из JSON данных."
    )
    parser.add_argument("--data", required=True, help="Путь к report_data.json")
    parser.add_argument("--style", required=True, help="Путь к style_config.json")
    parser.add_argument(
        "--output", default="generated_report.pptx", help="Путь к выходному .pptx файлу"
    )

    args = parser.parse_args()

    # Загрузка данных
    report_data = load_json(Path(args.data))
    style_data = load_json(Path(args.style))
    text_config = parse_text_config(style_data)

    # Генерация отчёта
    report = ReportGenerator(report_data, text_config=text_config)
    report.build()
    report.save(args.output)

    print(f"[✓] Отчёт успешно создан: {args.output}")


if __name__ == "__main__":
    main()
