# -*- coding: utf-8 -*-
__title__ = "RoomWarning"
__author__ = "JK_Sim"
__doc__ = """Version = 1.1
Date    = 03.11.2025
_____________________________________________________________________
Description:
Dismiss Multiple Rooms Warning and show alert when user clicks
_____________________________________________________________________
How-to:
-> Always ON on startup (registers silently)
-> Alert shown on manual button click
_____________________________________________________________________
"""

from pyrevit import forms
from pyrevit.revit import HOST_APP
import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FailureProcessingResult
from Autodesk.Revit.DB.Events import FailuresProcessingEventArgs
from System import EventHandler

# Keep reference to prevent garbage collection
event_handler = None

def handle_failures(sender, args):
    fa = args.GetFailuresAccessor()
    failures = fa.GetFailureMessages()
    to_delete = []

    for fm in failures:
        msg = fm.GetDescriptionText()
        if ("Multiple Rooms are in the same enclosed region" in msg or
            "Room Tag is outside of its Room" in msg):
            to_delete.append(fm)

    for fm in to_delete:
        fa.DeleteWarning(fm)

    args.SetProcessingResult(FailureProcessingResult.Continue)

def register_handler():
    """Register failure handler silently"""
    global event_handler
    app = HOST_APP.app

    if event_handler is None:
        event_handler = EventHandler[FailuresProcessingEventArgs](handle_failures)
        app.FailuresProcessing += event_handler

def enable_room_warning_handler():
    """Used when user clicks the button: shows alert and ensures handler is on"""
    register_handler()
    forms.alert("Room warning dismissal is ON", title="Dismiss Warning")

# Called by pyRevit when extension loads
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    register_handler()

# Called when user clicks the button manually
enable_room_warning_handler()
