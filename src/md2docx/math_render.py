from __future__ import annotations

from io import BytesIO

import matplotlib as mpl
from matplotlib import mathtext
from matplotlib.font_manager import FontProperties


def render_latex(expression: str, *, font_size: float = 16, dpi: int = 360) -> BytesIO:
    """Render TeX-like mathematics with a Times-compatible STIX math face."""
    expression = expression.strip()
    if not expression:
        raise ValueError("empty LaTeX expression")
    wrapped = expression if expression.startswith("$") else f"${expression}$"
    output = BytesIO()
    with mpl.rc_context(
        {
            "mathtext.fontset": "stix",
            "font.family": "STIXGeneral",
            "text.color": "black",
        }
    ):
        mathtext.math_to_image(
            wrapped,
            output,
            prop=FontProperties(family="STIXGeneral", size=font_size),
            dpi=dpi,
            format="png",
            color="black",
        )
    output.seek(0)
    return output
