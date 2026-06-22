from __future__ import annotations

from io import BytesIO

import matplotlib as mpl
from matplotlib import mathtext
from matplotlib.font_manager import FontProperties


def render_latex(
    expression: str,
    *,
    font_size: float = 16,
    font_family: str = "STIXGeneral",
    color: str = "black",
    dpi: int = 360,
) -> BytesIO:
    expression = expression.strip()
    if not expression:
        raise ValueError("empty LaTeX expression")
    wrapped = expression if expression.startswith("$") else f"${expression}$"
    output = BytesIO()
    with mpl.rc_context(
        {
            "mathtext.fontset": "stix",
            "font.family": font_family,
            "text.color": color,
        }
    ):
        mathtext.math_to_image(
            wrapped,
            output,
            prop=FontProperties(family=font_family, size=font_size),
            dpi=dpi,
            format="png",
            color=color,
        )
    output.seek(0)
    return output
