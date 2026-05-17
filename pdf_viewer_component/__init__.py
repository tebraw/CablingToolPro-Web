from pathlib import Path
import streamlit.components.v1 as components

_component_func = components.declare_component(
    "pdf_viewer",
    path=str(Path(__file__).parent),
)


def pdf_viewer_widget(img_b64, zoom, tx, ty, terms, key=None):
    """
    Renders the PDF viewer with pan/zoom and right-click → add-position menu.

    Parameters
    ----------
    img_b64 : str
        Base64-encoded PNG of the current page.
    zoom : float
        Current zoom level (passed through; not used by JS directly).
    tx, ty : float
        Initial translate offsets (auto-centering happens in JS on first load).
    terms : list[str]
        Search terms to show in the context menu.

    Returns
    -------
    dict | None
        {"action": "add_position", "term": <str>, "pdf_x": <float>, "pdf_y": <float>,
         "img_w": <int>, "img_h": <int>}
        or None if no interaction.
    """
    return _component_func(
        img_b64=img_b64,
        zoom=zoom,
        tx=tx,
        ty=ty,
        terms=terms,
        key=key,
        default=None,
    )
