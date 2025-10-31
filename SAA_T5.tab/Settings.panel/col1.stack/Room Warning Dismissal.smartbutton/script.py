# -*- coding: utf-8 -*-
__title__ = "RoomWarning"
__author__ = "JK_Sim"
__doc__ = """Version = 1.0
Date    = 30.10.2025
_____________________________________________________________________
Description:

Dismiss Multiple Rooms Warning
_____________________________________________________________________
How-to:

-> Always ON on startup
_____________________________________________________________________
"""

from pyrevit import forms
from pyrevit.revit import HOST_APP
import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FailureProcessingResult
from Autodesk.Revit.DB.Events import FailuresProcessingEventArgs
from System import EventHandler

# pyRevit expects this for .smartbutton, even if empty
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    pass


# Keep reference to prevent garbage collection
event_handler = None


def handle_failures(sender, args):
    """Event handler that dismisses specified warnings"""
    fa = args.GetFailuresAccessor()
    failures = fa.GetFailureMessages()
    to_delete = []

    for fm in failures:
        desc = fm.GetDescriptionText()
        if ("Multiple Rooms are in the same enclosed region" in desc or
            "Room Tag is outside of its Room" in desc):
            to_delete.append(fm)

    for fm in to_delete:
        fa.DeleteWarning(fm)

    args.SetProcessingResult(FailureProcessingResult.Continue)


def enable_room_warning_handler():
    global event_handler
    app = HOST_APP.app

    if event_handler is None:
        event_handler = EventHandler[FailuresProcessingEventArgs](handle_failures)
        app.FailuresProcessing += event_handler

    forms.alert("Room warning dismissal is ON", title="Dismiss Warning")


if __name__ == "__main__":
    enable_room_warning_handler()
