import math
import os

from PIL import Image

from .utils.text_style import apply_text_style
from .utils.units import mm


def add_image_slide(self, data):
    MAX_IMAGES_PER_SLIDE = 4
    images = data.get("images", [])

    if not images:
        return

    total_slides = math.ceil(len(images) / MAX_IMAGES_PER_SLIDE)

    for slide_idx in range(total_slides):
        layout = self.prs.slide_layouts[6]  # Blank slide
        slide = self.prs.slides.add_slide(layout)

        slide_width = self.prs.slide_width
        slide_height = self.prs.slide_height

        # === Заголовок ===
        title_height = mm(15)
        title_shape = slide.shapes.add_textbox(
            mm(10), mm(5), slide_width - mm(20), title_height
        )
        title_tf = title_shape.text_frame
        base_title = data.get("title", "")
        if total_slides > 1:
            title_tf.text = f"{base_title} (Часть {slide_idx + 1}/{total_slides})"
        else:
            title_tf.text = base_title

        for paragraph in title_tf.paragraphs:
            apply_text_style(paragraph, self.text_config.title)

        # === Параметры сетки ===
        start_idx = slide_idx * MAX_IMAGES_PER_SLIDE
        images_chunk = images[start_idx : start_idx + MAX_IMAGES_PER_SLIDE]
        chunk_len = len(images_chunk)

        cols = 2
        rows = math.ceil(chunk_len / cols)

        margin_x, margin_y = mm(10), title_height + mm(10)
        gap_x, gap_y = mm(10), mm(10)
        caption_height = mm(6)
        bottom_margin = mm(5)

        available_width = slide_width - 2 * margin_x - gap_x * (cols - 1)
        available_height = slide_height - margin_y - bottom_margin - gap_y * (rows - 1)

        cell_w = available_width / cols
        cell_h = available_height / rows

        # Учитываем подпись: максимум под картинку
        img_max_w = cell_w
        img_max_h = cell_h - caption_height

        for idx, image in enumerate(images_chunk):
            path = image.get("path")
            caption = image.get("caption", "")
            if not os.path.exists(path):
                continue

            col = idx % cols
            row = idx // cols

            cell_left = margin_x + col * (cell_w + gap_x)
            cell_top = margin_y + row * (cell_h + gap_y)

            # === Сохраняем пропорции изображения ===
            with Image.open(path) as img:
                orig_w, orig_h = img.size
                aspect_ratio = orig_w / orig_h

            target_w = img_max_w
            target_h = img_max_w / aspect_ratio

            if target_h > img_max_h:
                target_h = img_max_h
                target_w = img_max_h * aspect_ratio

            img_left = cell_left + (cell_w - target_w) / 2
            img_top = cell_top

            slide.shapes.add_picture(
                path, img_left, img_top, width=target_w, height=target_h
            )

            # === Подпись ===
            caption_top = img_top + target_h
            caption_box = slide.shapes.add_textbox(
                cell_left, caption_top, cell_w, caption_height
            )
            caption_tf = caption_box.text_frame
            caption_tf.text = caption
            for paragraph in caption_tf.paragraphs:
                apply_text_style(paragraph, self.text_config.body)
