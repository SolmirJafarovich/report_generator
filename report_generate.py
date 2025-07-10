import argparse
import json
from pathlib import Path

from src.report_generator.pptx_exporter import ReportGenerator
from src.report_generator.utils.pdf_converter import convert_pptx_to_pdf
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
    parser.add_argument(
        "--style", default="style_config.json", help="Путь к style_config.json"
    )
    parser.add_argument(
        "--output", default="generated_report.pptx", help="Путь к выходному .pptx файлу"
    )

    args = parser.parse_args()
    output_path = Path(args.output)

    # Загрузка данных
    report_data = load_json(Path(args.data))
    style_data = load_json(Path(args.style))
    text_config = parse_text_config(style_data)

    # Генерация отчёта
    report = ReportGenerator(report_data, text_config=text_config)
    report.build()

    if output_path.suffix.lower() == ".pdf":
        pptx_temp = output_path.with_suffix(".pptx")
        report.save(pptx_temp)
        success = convert_pptx_to_pdf(pptx_temp, output_path)

        if success:
            pptx_temp.unlink(missing_ok=True)
            print(f"[✓] PDF-отчёт успешно создан: {output_path}")
        else:
            print(
                f"[✗] Не удалось конвертировать PPTX в PDF. Файл PPTX сохранён: {pptx_temp}"
            )
    else:
        report.save(output_path)
        print(f"[✓] PPTX-отчёт успешно создан: {output_path}")

    print(f"[✓] Отчёт успешно создан: {args.output}")


if __name__ == "__main__":
    main()
