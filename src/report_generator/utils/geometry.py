def get_placeholder_geometry(prs, layout_idx=1, placeholder_idx=1):
    """Возвращает (left, top, width, height) placeholder'а слайдового layout'а без создания слайда."""
    layout = prs.slide_layouts[layout_idx]
    shape = layout.placeholders[placeholder_idx]
    return shape.left, shape.top, shape.width, shape.height
