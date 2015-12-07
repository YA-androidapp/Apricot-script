# -*- coding: utf-8 -*-
# Softalk.py
# Copyright (c) YA-androidapp All rights reserved.
# Required: 
#  SofTalk		( http://www35.atwiki.jp/softalk/pages/1.html )

import clr
clr.AddReferenceByPartialName("mscorlib")
clr.AddReferenceByPartialName("System")
clr.AddReferenceByPartialName("WindowsBase")
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName("Apricot")

from System import Object, Byte, Convert
from System.Collections.Generic import List
from System.IO import BufferedStream, BinaryWriter
from System.Diagnostics import Process
from System.Text import Encoding
from System.Threading.Tasks import Task, TaskCreationOptions
from System.Windows import Application, Window, DependencyPropertyChangedEventArgs
from Apricot import Balloon, Message, Script

import System
from System.Diagnostics import Process

def speech(text):
	def onSpeech():
		try:
			Process.Start("C:\\Program Files (x86)\\softalk\\SofTalk.exe","/W:"+text)
			Process.Start("C:\\Program Files (x86)\\softalk\\SofTalk.exe","/close")
			return None
		except:
			return None

	Task.Factory.StartNew(onSpeech, TaskCreationOptions.LongRunning)

def onIsVisibleChanged(s, e):
	if e.NewValue == True and s.Messages.Count > 0:
		speech(s.Messages[s.Messages.Count - 1].Text)

def onStart(s, e):
	global balloonList
	
	tempList = List[Balloon]()

	for window in Application.Current.Windows:
		if clr.GetClrType(Balloon).IsInstanceOfType(window):
			if not balloonList.Contains(window):
				window.IsVisibleChanged += onIsVisibleChanged
				balloonList.Add(window)
				
			tempList.Add(window)
		
	balloonList.Clear()
	balloonList.AddRange(tempList)

balloonList = List[Balloon]()
Script.Instance.Start += onStart