# -*- coding: utf-8 -*-
# Rate.py
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

# http://www.json.org/
class JsonDecoder(Object):
	TOKEN_NONE = 0
	TOKEN_CURLY_OPEN = 1
	TOKEN_CURLY_CLOSE = 2
	TOKEN_SQUARED_OPEN = 3
	TOKEN_SQUARED_CLOSE = 4
	TOKEN_COLON = 5
	TOKEN_COMMA = 6
	TOKEN_STRING = 7
	TOKEN_NUMBER = 8
	TOKEN_TRUE = 9
	TOKEN_FALSE = 10
	TOKEN_NULL = 11

	@staticmethod
	def decode(json):
		if json is not None:
			charArray = json.ToCharArray()
			index = 0
			value, index, success = JsonDecoder.parseValue(charArray, index, True)

			return value

		return None

	@staticmethod
	def parseObject(json, index, success):
		dictionary = Dictionary[String, Object]()
		token = -1
		i, index = JsonDecoder.nextToken(json, index) # {
		done = False

		while not done:
			token = JsonDecoder.lookAhead(json, index)

			if token == JsonDecoder.TOKEN_NONE:
				return None, index, False

			elif token == JsonDecoder.TOKEN_COMMA:
				i, index = JsonDecoder.nextToken(json, index)

			elif token == JsonDecoder.TOKEN_CURLY_CLOSE:
				i, index = JsonDecoder.nextToken(json, index)

				return dictionary, index, success

			else:
				name, index, success = JsonDecoder.parseString(json, index, success)

				if not success:
					return None, index, success

				token, index = JsonDecoder.nextToken(json, index)

				if token != JsonDecoder.TOKEN_COLON:
					return None, index, False

				value, index, success = JsonDecoder.parseValue(json, index, success)

				if not success:
					return None, index, success

				dictionary.Add(name, value)

		return dictionary, index, success

	@staticmethod
	def parseArray(json, index, success):
		list = List[Object]()
		i, index = JsonDecoder.nextToken(json, index) # [
		done = False

		while not done:
			token = JsonDecoder.lookAhead(json, index)

			if token == JsonDecoder.TOKEN_NONE:
				return None, index, False

			elif token == JsonDecoder.TOKEN_COMMA:
				i, index = JsonDecoder.nextToken(json, index)

			elif token == JsonDecoder.TOKEN_SQUARED_CLOSE:
				i, index = JsonDecoder.nextToken(json, index)

				break

			else:
				value, index, success = JsonDecoder.parseValue(json, index, success)

				if not success:
					return None, index, success

				list.Add(value)

		return list.ToArray(), index, success

	@staticmethod
	def parseValue(json, index, success):
		i = JsonDecoder.lookAhead(json, index)

		if i == JsonDecoder.TOKEN_STRING:
			return JsonDecoder.parseString(json, index, success)

		elif i == JsonDecoder.TOKEN_NUMBER:
			return JsonDecoder.parseNumber(json, index, success)

		elif i == JsonDecoder.TOKEN_CURLY_OPEN:
			return JsonDecoder.parseObject(json, index, success)

		elif i == JsonDecoder.TOKEN_SQUARED_OPEN:
			return JsonDecoder.parseArray(json, index, success)

		elif i == JsonDecoder.TOKEN_TRUE:
			j, index = JsonDecoder.nextToken(json, index)

			return True, index, success

		elif i == JsonDecoder.TOKEN_FALSE:
			j, index = JsonDecoder.nextToken(json, index)

			return False, index, success

		elif i == JsonDecoder.TOKEN_NULL:
			j, index = JsonDecoder.nextToken(json, index)

			return None, index, success

		return None, index, False

	@staticmethod
	def parseString(json, index, success):
		s = StringBuilder()
		index = JsonDecoder.skipWhitespace(json, index)
		c = json[index] # "
		index += 1
		complete = False
		
		while not complete:
			if index == json.Length:
				break

			c = json[index]
			index += 1

			if c == '"':
				complete = True
				break

			elif c == '\\':
				if index == json.Length:
					break

				c = json[index]
				index += 1

				if c == '"':
					s.Append('"')
				elif c == '\\':
					s.Append('\\')
				elif c == '/':
					s.Append('/')
				elif c == 'b':
					s.Append('\b')
				elif c == 'f':
					s.Append('\f')
				elif c == 'n':
					s.Append('\n')
				elif c == 'r':
					s.Append('\r')
				elif c == 't':
					s.Append('\t')
				elif c == 'u':
					remainingLength = json.Length - index

					if remainingLength >= 4:
						sb = StringBuilder()
						
						for i in range(4):
							sb.Append(json[index + i])

						success, codePoint = UInt32.TryParse(sb.ToString(), NumberStyles.HexNumber, CultureInfo.InvariantCulture)
						
						if not success:
							return String.Empty, index, success

						s.Append(Encoding.UTF32.GetString(BitConverter.GetBytes(codePoint)))
						index += 4

					else:
						break

			else:
				s.Append(c)

		if not complete:
			return None, index, False

		return s.ToString(), index, success

	@staticmethod
	def parseNumber(json, index, success):
		index = JsonDecoder.skipWhitespace(json, index)
		lastIndex = JsonDecoder.getLastIndexOfNumber(json, index)
		charLength = (lastIndex - index) + 1

		sb = StringBuilder()
						
		for i in range(charLength):
			sb.Append(json[index + i])

		success, number = Double.TryParse(sb.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture)
		index = lastIndex + 1
		
		return number, index, success

	@staticmethod
	def getLastIndexOfNumber(json, index):
		lastIndex = index

		while lastIndex < json.Length:
			if "0123456789+-.eE".IndexOf(json[lastIndex]) == -1:
				break

			lastIndex += 1

		return lastIndex - 1

	@staticmethod
	def skipWhitespace(json, index):
		while index < json.Length:
			if " \t\n\r".IndexOf(json[index]) == -1:
				break

			index += 1

		return index

	@staticmethod
	def lookAhead(json, index):
		i, j = JsonDecoder.nextToken(json, index)
		
		return i

	@staticmethod
	def nextToken(json, index):
		index = JsonDecoder.skipWhitespace(json, index)

		if index == json.Length:
			return JsonDecoder.TOKEN_NONE, index

		c = json[index]
		index += 1

		if c == '{':
			return JsonDecoder.TOKEN_CURLY_OPEN, index

		elif c == '}':
			return JsonDecoder.TOKEN_CURLY_CLOSE, index

		elif c == '[':
			return JsonDecoder.TOKEN_SQUARED_OPEN, index

		elif c == ']':
			return JsonDecoder.TOKEN_SQUARED_CLOSE, index

		elif c == ',':
			return JsonDecoder.TOKEN_COMMA, index

		elif c == '"':
			return JsonDecoder.TOKEN_STRING, index

		elif c == '0' or c == '1' or c == '2' or c == '3' or c == '4' or c == '5' or c == '6' or c == '7' or c == '8' or c == '9' or c == '-':
			return JsonDecoder.TOKEN_NUMBER, index

		elif c == ':':
			return JsonDecoder.TOKEN_COLON, index

		index -= 1
		remainingLength = json.Length - index

		# false
		if remainingLength >= 5:
			if json[index] == 'f' and json[index + 1] == 'a' and json[index + 2] == 'l' and json[index + 3] == 's' and json[index + 4] == 'e':
				index += 5

				return JsonDecoder.TOKEN_FALSE, index

		# true
		if remainingLength >= 4:
			if json[index] == 't' and json[index + 1] == 'r' and json[index + 2] == 'u' and json[index + 3] == 'e':
				index += 4

				return JsonDecoder.TOKEN_TRUE, index

		# null
		if remainingLength >= 4:
			if json[index] == 'n' and json[index + 1] == 'u' and json[index + 2] == 'l' and json[index + 3] == 'l':
				index += 4

				return JsonDecoder.TOKEN_NULL, index

		return JsonDecoder.TOKEN_NONE, index

def onTick(timer, e):
	def onUpdate():
		entryList = List[Entry]()

		if NetworkInterface.GetIsNetworkAvailable():
			# f = open('text.txt', 'w')
			try:
				request = WebRequest.Create(Uri("http://www.gaitameonline.com/rateaj/getrate"))
				response = None
				stream = None
				streamReader = None

				try:
					response = request.GetResponse()
					stream = response.GetResponseStream()
					streamReader = StreamReader(stream)
					r = streamReader.ReadToEnd()
					jsonDictionary = JsonDecoder.decode(r)
					# f.write(r)
					# f.write("\n\n\n")

					if jsonDictionary is not None and clr.GetClrType(Dictionary[String, Object]).IsInstanceOfType(jsonDictionary) and jsonDictionary.ContainsKey("quotes") and jsonDictionary["quotes"] is not None and clr.GetClrType(Array).IsInstanceOfType(jsonDictionary["quotes"]):
						for obj in jsonDictionary["quotes"]:
							ask = None
							bid = None
							currencyPairCode = None
							entry = Entry()

							if obj is not None and clr.GetClrType(Dictionary[String, Object]).IsInstanceOfType(obj):
								currencyPairCode = obj["currencyPairCode"]
								if (("EUR" in currencyPairCode) or ("GBP" in currencyPairCode) or ("USD" in currencyPairCode))and ("JPY" in currencyPairCode):
									currencyPairCode = currencyPairCode.replace("EUR", u"ユーロ").replace("GBP", u"ポンド").replace("USD", u"ドル").replace("JPY", "")
									ask = obj["ask"]
									bid = obj["bid"]
									val = String.Format("{0}: {1} / {2}", currencyPairCode, ask, bid)
									# f.write(val)
									# f.write("\n")

									entry.Title = val
									entry.Description = val
									entryList.Add(entry)
									# MessageBox.Show(val, currencyPairCode)

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

		# f.close()

		return RateList

	def onCompleted(task):
		if entryList.Count > 0:
			for entry in entryList:
				termList = getTermList(dictionary, entry.Title)
			
				if termList.Count > 0:
					sequenceList = List[Sequence]()

					for sequence in Script.Instance.Sequences:
						if sequence.Name.Equals("Activate"):
							sequenceList.Add(sequence)
					
					if Script.Instance.TryEnqueue(Script.Instance.Prepare(sequenceList, None, termList)):
						break

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
