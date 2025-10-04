"""
GUI Dialogs Package
"""

from gui.dialogs.settings_dialog import SettingsDialog

# Try to import optional dialogs
try:
    from gui.dialogs.appendix_edit_dialog import AppendixEditDialog
except ImportError:
    AppendixEditDialog = None

try:
    from gui.dialogs.preview_dialog import PreviewDialog
except ImportError:
    PreviewDialog = None

__all__ = ['SettingsDialog', 'AppendixEditDialog', 'PreviewDialog']