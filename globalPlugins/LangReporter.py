# Copyright (C) 2020  Alexander Linkov <kvark128@yandex.ru>
# Most of the code is taken from the NVDAHelper module and belongs to its authors

import globalPluginHandler
import api
import speech
import config
import sayAllHandler
import winUser
import addonHandler
import NVDAObjects.window
import ui
import queueHandler
import NVDAHelper
from NVDAObjects.UIA import UIA
from logHandler import log
from gui import SettingsPanel, NVDASettingsDialog, guiHelper

import wx
import winreg
from ctypes import *
from ctypes.wintypes import HKEY

addonHandler.initTranslation()

# Local information constants for obtaining of input language
# https://docs.microsoft.com/en-us/windows/win32/intl/locale-information-constants
LOCALE_SLANGUAGE = 0x00000002
LOCALE_SABBREVLANGNAME = 0x00000003
LOCALE_SNATIVELANGUAGENAME = 0x00000004
LOCALE_SLOCALIZEDLANGUAGENAME = 0x0000006f
LOCALE_SISO639LANGNAME = 0x00000059
LOCALE_SISO639LANGNAME2 = 0x00000067

LOCAL_CONSTANTS_DESCRIPTIONS = [
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

# Function for extracting a description of the keyboard layout from the registry by its hex code
def _lookupKeyboardLayoutNameWithHexString(layoutString):
	buf=create_unicode_buffer(1024)
	bufSize=c_int(2048)
	key=HKEY()
	if windll.advapi32.RegOpenKeyExW(winreg.HKEY_LOCAL_MACHINE,u"SYSTEM\\CurrentControlSet\\Control\\Keyboard Layouts\\"+ layoutString,0,winreg.KEY_QUERY_VALUE,byref(key))==0:
		try:
			if windll.advapi32.RegQueryValueExW(key,u"Layout Display Name",0,None,buf,byref(bufSize))==0:
				windll.shlwapi.SHLoadIndirectString(buf.value,buf,1023,None)
				return buf.value
			if windll.advapi32.RegQueryValueExW(key,u"Layout Text",0,None,buf,byref(bufSize))==0:
				return buf.value
		finally:
			windll.advapi32.RegCloseKey(key)

@WINFUNCTYPE(c_long,c_long,c_ulong,c_wchar_p)
def _nvdaControllerInternal_inputLangChangeNotify(threadID, hkl, layoutString):
	lastLanguageID = NVDAHelper.lastLanguageID
	lastLayoutString = NVDAHelper.lastLayoutString
	languageID = winUser.LOWORD(hkl)
	# Simple case where there is no change
	if languageID == lastLanguageID and layoutString == lastLayoutString:
		return 0
	focus = api.getFocusObject()
	# This callback can be called before NVDa is fully initialized
	# So also handle focus object being None as well as checking for sleepMode
	if not focus or focus.sleepMode:
		return 0
	# Generally we should not allow input lang changes from threads that are not focused.
	# But threadIDs for console windows are always wrong so don't ignore for those.
	if not isinstance(focus, NVDAObjects.window.Window) or (threadID != focus.windowThreadID and focus.windowClassName != "ConsoleWindowClass"):
		return 0
	# Never announce changes while in sayAll (#1676)
	if sayAllHandler.isRunning():
		return 0
	buf = create_unicode_buffer(1024)
	res = windll.kernel32.GetLocaleInfoW(languageID, config.conf["LangReporter"]["languagePresentation"], buf, 1024)
	# Translators: the label for an unknown language when switching input methods.
	inputLanguageName = buf.value if res else _("unknown language")
	layoutStringCodes = []
	inputMethodName = None
	# LayoutString can either be a real input method name, a hex string for an input method name in the registry, or an empty string.
	# If it is a real input method name, then it is used as is.
	# If it is a hex string or it is empty, then the method name is looked up by trying:
	# The full hex string, the hkl as a hex string, the low word of the hex string or hkl, the high word of the hex string or hkl.
	if layoutString:
		try:
			int(layoutString, 16)
			layoutStringCodes.append(layoutString)
		except ValueError:
			inputMethodName = layoutString
	if not inputMethodName:
		layoutStringCodes.insert(0,hex(hkl)[2:].rstrip('L').upper().rjust(8,'0'))
		for stringCode in list(layoutStringCodes):
			layoutStringCodes.append(stringCode[4:].rjust(8,'0'))
			if stringCode[0]<'D':
				layoutStringCodes.append(stringCode[0:4].rjust(8,'0'))
		for stringCode in layoutStringCodes:
			inputMethodName = _lookupKeyboardLayoutNameWithHexString(stringCode)
			if inputMethodName: break
	if not inputMethodName:
		log.debugWarning("Could not find layout name for keyboard layout, reporting as unknown") 
		# Translators: The label for an unknown input method when switching input methods. 
		inputMethodName = _("unknown input method")
	# Remove the language name if it is in the input method name.
	if ' - ' in inputMethodName:
		inputMethodName = "".join(inputMethodName.split(' - ')[1:])
	# Include the language only if it changed.
	if languageID != lastLanguageID:
		if config.conf["LangReporter"]["reportLayout"]:
			msg = _("{language} - {layout}").format(language=inputLanguageName, layout=inputMethodName)
		else:
			msg = inputLanguageName
	else:
		msg = inputMethodName
	NVDAHelper.lastLanguageID = languageID
	NVDAHelper.lastLayoutString = layoutString
	queueHandler.queueFunction(queueHandler.eventQueue, ui.message, msg)
	return 0

class InputSwitch(UIA):

	def _get_shouldAllowUIAFocusEvent(self):
		return False

	def event_UIA_elementSelected(self):
		name = self.name
		if config.conf["LangReporter"]["reportLanguageSwitchingBar"] and name:
			speech.cancelSpeech()
			ui.message(name)

class AddonSettingsPanel(SettingsPanel):
	title = _("Input language and layout")

	def makeSettings(self, settingsSizer):
		sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		presentationChoices = [f[1] for f in LOCAL_CONSTANTS_DESCRIPTIONS]
		self.languagePresentationChoice = sHelper.addLabeledControl(_("Input language &presentation:"), wx.Choice, choices=presentationChoices)
		curPresentation = config.conf["LangReporter"]["languagePresentation"]
		try:
			index = [f[0] for f in LOCAL_CONSTANTS_DESCRIPTIONS].index(curPresentation)
			self.languagePresentationChoice.SetSelection(index)
		except Exception:
			log.error(f"invalid language presentation: {curPresentation}")

		self.reportLayoutCheckBox = sHelper.addItem(wx.CheckBox(self, label=_("Report &layout when language switching")))
		self.reportLayoutCheckBox.SetValue(config.conf["LangReporter"]["reportLayout"])

		self.reportLanguageSwitchingBarCheckBox = sHelper.addItem(wx.CheckBox(self, label=_("Report language &switching bar")))
		self.reportLanguageSwitchingBarCheckBox.SetValue(config.conf["LangReporter"]["reportLanguageSwitchingBar"])

	def postInit(self):
		self.languagePresentationChoice.SetFocus()

	def onSave(self):
		config.conf["LangReporter"]["languagePresentation"] = LOCAL_CONSTANTS_DESCRIPTIONS[self.languagePresentationChoice.GetSelection()][0]
		config.conf["LangReporter"]["reportLayout"] = self.reportLayoutCheckBox.GetValue()
		config.conf["LangReporter"]["reportLanguageSwitchingBar"] = self.reportLanguageSwitchingBarCheckBox.GetValue()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

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
