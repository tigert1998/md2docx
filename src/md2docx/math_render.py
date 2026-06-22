from __future__ import annotations

from io import BytesIO

import matplotlib as mpl
from matplotlib import mathtext
from matplotlib.font_manager import FontProperties


def render_latex(
    expression: str,
    *,
    font_size: float = 16,
    fontset: str = "stix",
    color: str = "black",
    dpi: int = 360,
) -> BytesIO:
    expression = expression.strip()
    if not expression:
        raise ValueError("empty LaTeX expression")
    wrapped = expression if expression.startswith("$") else f"${expression}$"
    output = BytesIO()
    try:
        with mpl.rc_context({"mathtext.fontset": fontset}):
            mathtext.math_to_image(
                wrapped,
                output,
                prop=FontProperties(size=font_size),
                dpi=dpi,
                format="png",
                color=color,
            )
    except ValueError as exc:
        if "mathtext.fontset" in str(exc):
            raise ValueError(f"unsupported math fontset: {fontset}") from exc
        raise
    output.seek(0)
    return output
