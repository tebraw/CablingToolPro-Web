from pathlib import Path
import streamlit.components.v1 as components

_component_func = components.declare_component(
    "cable_list",
    path=str(Path(__file__).parent),
)


def cable_list_widget(items, key=None):
    """
    Renders a sortable, editable, deletable list of cable entries.

    Parameters
    ----------
    items : list[dict]
        Each dict must contain:
          - ukv_text  (str)  – the found text shown on the left (read-only)
          - label     (str)  – our label shown on the right (editable)
          - checked   (bool) – whether this entry is active
          - _gi       (int)  – global index into kabel_fields (pass-through)

    Returns
    -------
    list[dict] | None
        Updated items list (same structure) if user interacted, else None.
    """
    return _component_func(items=items, key=key, default=None)
