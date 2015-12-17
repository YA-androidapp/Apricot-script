# -*- coding: utf-8 -*-
# Nkave.py
# Copyright c Masaaki Kawata All rights reserved.
# Copyright (c) YA-androidapp All rights reserved.

import clr
clr.AddReferenceByPartialName("mscorlib")
clr.AddReferenceByPartialName("System")
clr.AddReferenceByPartialName("System.Core")
clr.AddReferenceByPartialName("System.Windows.Forms")
clr.AddReferenceByPartialName("WindowsBase")
clr.AddReferenceByPartialName("PresentationCore")
clr.AddReferenceByPartialName("PresentationFramework")
clr.AddReferenceByPartialName("Apricot")

from System import Object, Nullable, Byte, UInt32, Double, Char, String, Uri, DateTime, TimeSpan, Array, StringComparison, Convert, BitConverter, Math, Action#, DateTime
from System.Collections.Generic import List, Dictionary, KeyValuePair, HashSet
from System.IO import Stream, StreamReader, MemoryStream, FileStream, SeekOrigin, FileMode, FileAccess, FileShare
from System.Diagnostics import Trace
from System.Globalization import CultureInfo, NumberStyles
from System.Text import StringBuilder, Encoding
from System.Text.RegularExpressions import Regex, Match, RegexOptions
from System.Threading.Tasks import Task, TaskCreationOptions, TaskScheduler
from System.Linq import Enumerable
from System.Net import WebRequest, WebResponse
from System.Net.NetworkInformation import NetworkInterface
from System.Windows import Application, Window, WindowStartupLocation, WindowStyle, ResizeMode, SizeToContent, HorizontalAlignment, VerticalAlignment, Rect, PropertyPath
from System.Windows.Controls import ContentControl, Grid
from System.Windows.Forms import MessageBox
from System.Windows.Media import Color, Colors, Brushes, SolidColorBrush, TranslateTransform, ScaleTransform, RectangleGeometry, EllipseGeometry, ImageBrush, TileMode, BrushMappingMode, Stretch, AlignmentX, AlignmentY, DrawingGroup, DrawingContext, DrawingImage
from System.Windows.Media.Animation import Storyboard, Clock, ClockState, DoubleAnimation, SineEase, EasingMode
from System.Windows.Media.Effects import DropShadowEffect
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption, BitmapCreateOptions
from System.Windows.Shapes import Rectangle
from System.Windows.Threading import DispatcherTimer, DispatcherPriority
from Apricot import Script, Entry, Sequence
import re

def onTick(timer, e):
	def onUpdate():
		entryList = List[Entry]()

		if NetworkInterface.GetIsNetworkAvailable():
			try:
				request = WebRequest.Create(Uri("http://indexes.nikkei.co.jp/nkave"))
				response = None
				stream = None
				streamReader = None

				try:
					response = request.GetResponse()
					stream = response.GetResponseStream()
					streamReader = StreamReader(stream)
					r = streamReader.ReadToEnd()

					pattern1 = re.compile("top-nk225-value\">([,.0-9]+)[^0-9]")
					pattern2 = re.compile("top-nk225-differ\">([-+,.0-9]+)[^0-9]")
					m1 = pattern1.search(r)
					m2 = pattern2.search(r)
					if m1:
						entry = Entry()
						val = String.Format(u"日経平均株価 {0} ({1})", m1.group(1), m2.group(1))
						entry.Title = val
						entry.Description = val
						entryList.Add(entry)
						Script.Instance.Alert(entryList)

				finally:
					if streamReader is not None:
						streamReader.Close()

					if stream is not None:
						stream.Close()
			
					if response is not None:
						response.Close()

			except Exception, e:
				Trace.WriteLine(e.clsException.Message)
				Trace.WriteLine(e.clsException.StackTrace)

		return entryList

	def onCompleted(task):
		if task.Result.Key.Count > 0:
			sequenceList = List[Sequence]()

			for sequence in Script.Instance.Sequences:
				if sequence.Name.Equals("Activate"):
					sequenceList.Add(sequence)

			for s in task.Result.Key:
				Script.Instance.TryEnqueue(Script.Instance.Prepare(sequenceList, s))

	Task.Factory.StartNew(onUpdate, TaskCreationOptions.LongRunning).ContinueWith(onCompleted, TaskScheduler.FromCurrentSynchronizationContext())

	timer.Stop()
	timer.Interval = TimeSpan.FromMinutes(5)
	timer.Start()

def urlEncode(value):
	unreserved = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~"
	sb = StringBuilder()
	bytes = Encoding.UTF8.GetBytes(value)

	for b in bytes:
		if b < 0x80 and unreserved.IndexOf(Convert.ToChar(b)) != -1:
			sb.Append(Convert.ToChar(b))
		else:
			sb.Append('%' + String.Format("{0:X2}", Convert.ToInt32(b)))

	return sb.ToString()

def onStart(s, e):
	global timer
	timer.Start()

def onStop(s, e):
	global timer
	timer.Stop()

timer = DispatcherTimer(DispatcherPriority.Background)
timer.Tick += onTick
timer.Interval = TimeSpan.FromMilliseconds(30000)
Script.Instance.Start += onStart
Script.Instance.Stop += onStop
