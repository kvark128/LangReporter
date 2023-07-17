# Copyright (C) 2020 - 2023 Alexander Linkov <kvark128@yandex.ru>
# This file is covered by the GNU General Public License.
# See the file COPYING.txt for more details.
# Ukrainian Nazis and their accomplices are not allowed to use this plugin. Za pobedu!

import threading
import os
import winsound
import itertools
import functools
import wx
import winreg
from ctypes import *

import globalPluginHandler
import api
import globalCommands
import gui
import speech
import config
import globalVars
import winUser
import winVersion
import addonHandler
import ui
import queueHandler
import NVDAHelper
from speech import sayAll
from scriptHandler import script
from NVDAObjects.UIA import UIA
from NVDAObjects.window import Window
from logHandler import log
from gui import guiHelper
from gui.settingsDialogs import SettingsPanel, NVDASettingsDialog

addonHandler.initTranslation()

MODULE_DIR = os.path.dirname(__file__)
WARN_SND_PATH = os.path.join(MODULE_DIR, "warn.wav")

# Message posted to the window with focus to change the current keyboard layout
# https://docs.microsoft.com/en-us/windows/win32/winmsg/wm-inputlangchangerequest
WM_INPUTLANGCHANGEREQUEST = 0x0050

# Local information constants for obtaining of input language
# https://docs.microsoft.com/en-us/windows/win32/intl/locale-information-constants
LOCALE_SLANGUAGE = 0x00000002
LOCALE_SABBREVLANGNAME = 0x00000003
LOCALE_SNATIVELANGUAGENAME = 0x00000004
LOCALE_SLOCALIZEDLANGUAGENAME = 0x0000006f
LOCALE_SISO639LANGNAME = 0x00000059
LOCALE_SISO639LANGNAME2 = 0x00000067

LOCAL_CONSTANTS_LABELS = [
	(LOCALE_SLANGUAGE, _("Full localized name of the language")),
	(LOCALE_SABBREVLANGNAME, _("Abbreviated name of the language")),
	(LOCALE_SNATIVELANGUAGENAME, _("Native name of the language")),
	(LOCALE_SLOCALIZEDLANGUAGENAME, _("Localized name of the language")),
	(LOCALE_SISO639LANGNAME, _("Two-letter language name from ISO 639")),
	(LOCALE_SISO639LANGNAME2, _("Three-letter language name from ISO 639-2")),
]

# Utility function to point an exported function pointer in a dll  to a ctypes wrapped python function
def _setDllFuncPointer(dll, name, cfunc):
	cast(getattr(dll,name),POINTER(c_void_p)).contents.value=cast(cfunc,c_void_p).value

# Function for extracting a description of the keyboard layout from the registry by its klid
def getKeyboardLayoutDisplayName(klid):
	buf = create_unicode_buffer(1024)
	bufSize = c_int(2048)
	key = wintypes.HKEY()
	if windll.advapi32.RegOpenKeyExW(winreg.HKEY_LOCAL_MACHINE, rf"SYSTEM\CurrentControlSet\Control\Keyboard Layouts\{klid}", 0, winreg.KEY_QUERY_VALUE, byref(key)) == 0:
		try:
			if windll.advapi32.RegQueryValueExW(key, "Layout Display Name", 0, None, buf, byref(bufSize)) == 0:
				windll.shlwapi.SHLoadIndirectString(buf.value, buf, 1023, None)
				return buf.value
			if windll.advapi32.RegQueryValueExW(key, "Layout Text", 0, None, buf, byref(bufSize)) == 0:
				return buf.value
		finally:
			windll.advapi32.RegCloseKey(key)

# Returns KLID string same as GetKeyboardLayoutName, but for any HKL
# Based on https://github.com/dotnet/winforms/issues/4345#issuecomment-759161693
def getKLIDFromHKL(hkl):
	deviceID = winUser.HIWORD(hkl)
	if deviceID & 0xf000 != 0xf000: # deviceID not contains layoutID
		langID = winUser.LOWORD(hkl)
		if deviceID != 0:
			# deviceID overrides langID if set
			langID = deviceID
		return f"{langID:08x}"

	key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Keyboard Layouts")
	layoutID = deviceID & 0x0fff
	for i in itertools.count():
		try:
			klid = winreg.EnumKey(key, i)
		except OSError:
			return None
		try:
			klid_key = winreg.OpenKey(key, klid)
			data, data_type = winreg.QueryValueEx(klid_key, "Layout Id")
			if data_type == winreg.REG_SZ and int(data, 16) == layoutID:
				return klid
		except (FileNotFoundError, ValueError, TypeError):
			pass

def _lookupKeyboardLanguageName(lcid, lctype):
	buf = create_unicode_buffer(1024)
	res = windll.kernel32.GetLocaleInfoW(lcid, lctype, buf, 1024)
	# Translators: the label for an unknown language
	return buf.value if res else _("unknown language")

@functools.lru_cache(maxsize=None)
def _lookupKeyboardLayoutName(hkl):
	klid = getKLIDFromHKL(hkl)
	if klid is None:
		log.debugWarning("Could not find klid for hkl, reporting as unknown")
		# Translators: The label for an unknown input method
		return _("unknown input method")
	inputMethodName = getKeyboardLayoutDisplayName(klid)
	# Remove the language name if it is in the input method name.
	if ' - ' in inputMethodName:
		inputMethodName = "".join(inputMethodName.split(' - ')[1:])
	return inputMethodName

_inputSwitchingBarWillBeAnnounced = threading.Event()
_inputSwitchingLock = threading.Lock()
_lastFocusWhenLanguageSwitching = None
_last_hkl = 0

@WINFUNCTYPE(c_long,c_long,c_ulong,c_wchar_p)
def _nvdaControllerInternal_inputLangChangeNotify(threadID, hkl, layoutString):
	global _lastFocusWhenLanguageSwitching, _last_hkl
	languageID = winUser.LOWORD(hkl)
	languageSwitching = True
	with _inputSwitchingLock:
		if hkl == _last_hkl:
			# Simple case where there is no change
			return 0
		if languageID == winUser.LOWORD(_last_hkl):
			languageSwitching = False
	focus = api.getFocusObject()
	# This callback can be called before NVDa is fully initialized
	# So also handle focus object being None as well as checking for sleepMode
	if not focus or focus.sleepMode:
		return 0
	# Generally we should not allow input lang changes from threads that are not focused.
	# But threadIDs for console windows are always wrong so don't ignore for those.
	if not isinstance(focus, Window) or (threadID != focus.windowThreadID and focus.windowClassName != "ConsoleWindowClass"):
		return 0
	with _inputSwitchingLock:
		_lastFocusWhenLanguageSwitching = focus
		_last_hkl = hkl
	# Never announce changes while in sayAll (#1676)
	if sayAll.SayAllHandler.isRunning():
		return 0
	# Never announce changes if it has already been done in the language switching bar
	if _inputSwitchingBarWillBeAnnounced.isSet():
		_inputSwitchingBarWillBeAnnounced.clear()
		return 0
	inputMethodName = _lookupKeyboardLayoutName(hkl)
	# Include the language only if it changed.
	if languageSwitching:
		inputLanguageName = _lookupKeyboardLanguageName(languageID, config.conf["LangReporter"]["languagePresentation"])
		if config.conf["LangReporter"]["reportLayout"]:
			msg = _("{language} - {layout}").format(language=inputLanguageName, layout=inputMethodName)
		else:
			msg = inputLanguageName
	else:
		msg = inputMethodName
	queueHandler.queueFunction(queueHandler.eventQueue, ui.message, msg)
	return 0

class InputSwitch(UIA):

	def _get_shouldAllowUIAFocusEvent(self):
		return False

	def event_UIA_elementSelected(self):
		inputMethodName = self.name
		lWinIsPressed = bool(winUser.getAsyncKeyState(winUser.VK_LWIN) & 0x8000)
		rWinIsPressed = bool(winUser.getAsyncKeyState(winUser.VK_RWIN) & 0x8000)
		if inputMethodName and (lWinIsPressed or rWinIsPressed):
			# User switches layout with windows+space
			if config.conf["LangReporter"]["reportLanguageSwitchingBar"]:
				_inputSwitchingBarWillBeAnnounced.set()
				speech.cancelSpeech()
				ui.message(inputMethodName)

class AddonSettingsPanel(SettingsPanel):
	title = _("Input language and layout")

	def makeSettings(self, settingsSizer):
		sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		presentationChoices = [f[1] for f in LOCAL_CONSTANTS_LABELS]
		self.languagePresentationChoice = sHelper.addLabeledControl(_("Input language &presentation:"), wx.Choice, choices=presentationChoices)
		curPresentation = config.conf["LangReporter"]["languagePresentation"]
		try:
			index = [f[0] for f in LOCAL_CONSTANTS_LABELS].index(curPresentation)
			self.languagePresentationChoice.SetSelection(index)
		except Exception:
			log.error(f"invalid language presentation: {curPresentation}")

		self.reportLayoutCheckBox = sHelper.addItem(wx.CheckBox(self, label=_("Report &layout when language switching")))
		self.reportLayoutCheckBox.SetValue(config.conf["LangReporter"]["reportLayout"])

		self.reportLanguageSwitchingBarCheckBox = sHelper.addItem(wx.CheckBox(self, label=_("Report language &switching bar when pressing Windows+Space")))
		self.reportLanguageSwitchingBarCheckBox.SetValue(config.conf["LangReporter"]["reportLanguageSwitchingBar"])
		# Switching layouts via Windows+Space is available starting with Windows 8
		self.reportLanguageSwitchingBarCheckBox.Enable(winVersion.getWinVer() >= winVersion.WIN8)

	def postInit(self):
		self.languagePresentationChoice.SetFocus()

	def onSave(self):
		config.conf["LangReporter"]["languagePresentation"] = LOCAL_CONSTANTS_LABELS[self.languagePresentationChoice.GetSelection()][0]
		config.conf["LangReporter"]["reportLayout"] = self.reportLayoutCheckBox.GetValue()
		config.conf["LangReporter"]["reportLanguageSwitchingBar"] = self.reportLanguageSwitchingBarCheckBox.GetValue()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def __new__(cls, *args, **kwargs):
		size = windll.user32.GetKeyboardLayoutList(0, None)
		if size > 0:
			HKLList = (wintypes.HKL * size)()
			windll.user32.GetKeyboardLayoutList(size, HKLList)
			for hkl in HKLList:
				cls.addScriptForHKL(hkl)
		return super().__new__(cls, *args, **kwargs)

	@classmethod
	def addScriptForHKL(cls, hkl):
		script = lambda self, gesture: self._switchToLayout(hkl)
		funcName = script.__name__ = f"script_switchLayoutTo_{hkl:08x}"
		script.category = _("Keyboard layouts")
		inputLanguageName = _lookupKeyboardLanguageName(winUser.LOWORD(hkl), LOCALE_SLANGUAGE)
		inputMethodName = _lookupKeyboardLayoutName(hkl)
		script.__doc__ = _("Switches to {language} - {layout} layout").format(language=inputLanguageName, layout=inputMethodName)
		setattr(cls, funcName, script)

	def _switchToLayout(self, hkl):
		focus = api.getFocusObject()
		current_hkl = c_ulong(windll.User32.GetKeyboardLayout(focus.windowThreadID)).value
		if hkl == current_hkl:
			# Requested layout is already active
			if os.path.isfile(WARN_SND_PATH):
				winsound.PlaySound(WARN_SND_PATH, winsound.SND_ASYNC)
			return
		winUser.PostMessage(focus.windowHandle, WM_INPUTLANGCHANGEREQUEST, 0, hkl)

	def __init__(self):
		super().__init__()
		config.conf.spec["LangReporter"] = {
			"languagePresentation": f"integer(default={LOCALE_SLANGUAGE})",
			"reportLayout": "boolean(default=True)",
			"reportLanguageSwitchingBar": "boolean(default=True)",
		}
		NVDASettingsDialog.categoryClasses.append(AddonSettingsPanel)
		_setDllFuncPointer(NVDAHelper.localLib, "_nvdaControllerInternal_inputLangChangeNotify", _nvdaControllerInternal_inputLangChangeNotify)

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if isinstance(obj, UIA) and "InputSwitch.dll" in obj.UIAElement.cachedProviderDescription:
			clsList.insert(0, InputSwitch)

	def terminate(self):
		_setDllFuncPointer(NVDAHelper.localLib, "_nvdaControllerInternal_inputLangChangeNotify", NVDAHelper.nvdaControllerInternal_inputLangChangeNotify)
		NVDASettingsDialog.categoryClasses.remove(AddonSettingsPanel)

	def event_foreground(self, obj, nextHandler):
		# Different windows may have different input languages
		# We need to update last language information when switching between windows
		global _last_hkl
		with _inputSwitchingLock:
			_last_hkl = c_ulong(windll.User32.GetKeyboardLayout(obj.windowThreadID)).value
		nextHandler()

	def event_gainFocus(self, obj, nextHandler):
		# Starting with Windows 10 version 1903, switching the input language causes NVDA to create a fake focus event
		# We need to prevent the handling of such an event once
		global _lastFocusWhenLanguageSwitching
		with _inputSwitchingLock:
			lastFocus = _lastFocusWhenLanguageSwitching
			_lastFocusWhenLanguageSwitching = None
			if obj == lastFocus:
				return
		nextHandler()

	@script(
		description=_("Shows input language and layout presentation settings"),
		category=globalCommands.SCRCAT_CONFIG,
	)
	def script_activateInputLanguagePresentationDialog(self, gesture):
		if not globalVars.appArgs.secure:
			gui.mainFrame._popupSettingsDialog(NVDASettingsDialog, AddonSettingsPanel)
