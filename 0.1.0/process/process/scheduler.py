from typing import Union
from pathlib import Path
from win32com import client
from dataclasses import dataclass
from datetime import datetime as dt
from dateutil.parser import parse as dtparse


# [NOTE] This file is not complete.


@dataclass
class Calendar:

    start: Union[str, dt] 
    end: Union[str, dt] 
    repeat: str



class Task:

    
    TRIGGER_TYPE_DAILY = 2
    ACTION_TYPE_EXEC = 0

    
    def __init__(self, definition: object, name: str, desc: str, start: Union[str, dt] = None,
                 end: Union[str, dt] = None, repeat: str = None):
        # [TODO] inner folder on the basis of name?
        
        self.cal = Calendar(start, end, repeat)
        self.cal.schedule()
        
        # [NOTE] Quit must be called prior to del
        self.definition = definition 
        regInfo = self.definition.RegistrationInfo
        regInfo.Description = desc
        regInfo.Author = "Administrator"
        self.regInfo = regInfo
        
        settings = self.definition.Settings
        settings.Enabled = True
        settings.StartWhenAvailable = True
        settings.Hidden = False
        self.settings = settings

        triggers = self.definition.Triggers
        trigger = triggers.Create(self.TRIGGER_TYPE_DAILY)
        trigger.StartBoundary = ""#???
        trigger.EndBoundary = ""#??? is this necessary?
        trigger.DaysInterval = 1 #??? then why TriggerTypeDaily???
        trigger.Id = "DailyTriggerId" #??? I assume I make this up?
        trigger.Enabled = True
        self.trigger = trigger

        repetitionPattern = trigger.Repetition
        repetitionPattern.Duration = "PT4M" # seriously????
        repetitionPattern.Interval = "PT1M"
        
        # [TODO] parse out repeat into "PT4M"...
        # if repeat
        self.repeat = repetitionPattern

        
class Scheduler:

    def __init__(self):
        self.app = client.Dispatch('Schedule.Service')

    def __del__(self):
        self.app.Quit()
            
    def add_task(self, name: str, desc: str, exec, flags,
                  userID: str = None, password: str = None,
                  start = None, end = None, repeat = None):
        exec = Path(exec)
        if not exec.exists():
            raise FileNotFoundError
        
        definition = self.app.NewTask(0)
        task = Task(definition, name, desc, start, end, repeat)
        
        # and if not ActionTypeExec???
        action = task.definition.Actions.Create(Task.ActionTypeExec)
        action.Path = exec
        
        rootFolder = self.app.GetFolder("\\")
        userFlag = "Need to find the user flag vs admin flag"
        #rootFolder.RegisterTaskDefinition(None, name, task.definition,
        #(task.flags), userID, password, userFlag) 


# [NOTE] Placeholder until release.
Calendar = lambda: print("This feature is not yet available.")
Task = lambda: print("This feature is yet available.")
Scheduler = lambda: print("This feature is yet available.")