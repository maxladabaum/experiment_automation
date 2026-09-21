from types import SimpleNamespace
from unittest.mock import Mock

from gui.widgets import ScrollableFrame


class FakeWidget:
    def __init__(self, widget_class, children=None):
        self._widget_class = widget_class
        self._children = list(children or [])
        self.bindings = {}

    def winfo_class(self):
        return self._widget_class

    def winfo_children(self):
        return list(self._children)

    def bind(self, sequence, callback, add=None):
        self.bindings[sequence] = callback


def test_scrollable_frame_routes_combobox_wheel_without_changing_selection():
    combo = FakeWidget("TCombobox")
    entry = FakeWidget("TEntry")
    content = FakeWidget("TFrame", [FakeWidget("TFrame", [combo, entry])])
    frame = ScrollableFrame.__new__(ScrollableFrame)
    frame._canvas = Mock()

    frame._protect_descendant_comboboxes(content)

    assert "<MouseWheel>" in combo.bindings
    assert "<MouseWheel>" not in entry.bindings
    result = combo.bindings["<MouseWheel>"](SimpleNamespace(delta=-120))
    assert result == "break"
    frame._canvas.yview_scroll.assert_called_once_with(1, "units")
