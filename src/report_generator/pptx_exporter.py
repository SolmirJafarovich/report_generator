from pathlib import Path

from pptx import Presentation

from .slide_builder import SlideBuilder
from .utils.text_config import TextConfig, TextStyle


class ReportGenerator:
    def __init__(self, report_data: dict):
        self.data = report_data
        self.template_path = report_data.get("template_path")

        # === Загружаем Presentation ===
        self.prs = self._load_template()

        # === Загружаем TextConfig ===
        self.text_config = self._load_text_config()

        # === Создаём Builder с конфигом ===
        self.builder = SlideBuilder(self.prs, self.text_config)

    def _load_template(self):
        if self.template_path and Path(self.template_path).exists():
            return Presentation(self.template_path)
        return Presentation()  # пустая презентация, если шаблон не задан

    def _load_text_config(self):
        # Можно брать из self.data, если ты хочешь задавать конфиг через JSON/YAML
        title = self.data.get("text_config", {}).get("title", {})
        body = self.data.get("text_config", {}).get("body", {})
        table_header = self.data.get("text_config", {}).get("table_header", {})
        table_cell = self.data.get("text_config", {}).get("table_cell", {})

        return TextConfig(
            title=TextStyle(**title) if title else None,
            body=TextStyle(**body) if body else None,
            table_header=TextStyle(**table_header) if table_header else None,
            table_cell=TextStyle(**table_cell) if table_cell else None,
        )

    def build(self):
        for slide_data in self.data.get("slides", []):
            self.builder.add_slide(slide_data)

    def save(self, output_path: str = "output.pptx"):
        self.prs.save(output_path)
        print(f"[✓] Презентация сохранена: {output_path}")
