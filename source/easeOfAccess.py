# A part of NonVisual Desktop Access (NVDA)
# Copyright (C) 2014-2022 NV Access Limited
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

"""Utilities for working with the Windows Ease of Access Center.
"""

from typing import List, TYPE_CHECKING, Union
from logHandler import log
import winreg
import winUser
import winVersion


if TYPE_CHECKING:
	from winreg import HKEYType
	_KeyType = Union[HKEYType, int]


# Windows >= 8
canConfigTerminateOnDesktopSwitch: bool = winVersion.getWinVer() >= winVersion.WIN8

ROOT_KEY = r"Software\Microsoft\Windows NT\CurrentVersion\Accessibility"
APP_KEY_NAME = "nvda_nvda_v1"
APP_KEY_PATH = r"%s\ATs\%s" % (ROOT_KEY, APP_KEY_NAME)


def isRegistered() -> bool:
	try:
		winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, APP_KEY_PATH, 0,
			winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
		return True
	except FileNotFoundError:
		log.debug("Unable to find AT registry key")
	except WindowsError:
		log.error("Unable to open AT registry key", exc_info=True)
	return False


def notify(signal):
	if not isRegistered():
		return
	with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\AccessibilityTemp") as rkey:
		winreg.SetValueEx(rkey, APP_KEY_NAME, None, winreg.REG_DWORD, signal)
	keys = []
	# The user might be holding unwanted modifiers.
	for vk in winUser.VK_SHIFT, winUser.VK_CONTROL, winUser.VK_MENU:
		if winUser.getAsyncKeyState(vk) & 32768:
			keys.append((vk, False))
	keys.append((0x5B, True)) # leftWindows
	keys.append((0x55, True)) # u
	inputs = []
	# Release unwanted keys and press desired keys.
	for vk, desired in keys:
		input = winUser.Input(type=winUser.INPUT_KEYBOARD)
		input.ii.ki.wVk = vk
		if not desired:
			input.ii.ki.dwFlags = winUser.KEYEVENTF_KEYUP
		inputs.append(input)
	# Release desired keys and press unwanted keys.
	for vk, desired in reversed(keys):
		input = winUser.Input(type=winUser.INPUT_KEYBOARD)
		input.ii.ki.wVk = vk
		if desired:
			input.ii.ki.dwFlags = winUser.KEYEVENTF_KEYUP
		inputs.append(input)
	winUser.SendInput(inputs)


def willAutoStart(hkey: _KeyType) -> bool:
	return (APP_KEY_NAME in _getAutoStartConfiguration(hkey))


def _getAutoStartConfiguration(hkey: _KeyType) -> List[str]:
	try:
		k = winreg.OpenKey(hkey, ROOT_KEY, 0,
			winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	except FileNotFoundError:
		log.debug("Unable to find existing auto start registry key")
		return []
	except WindowsError:
		log.error("Unable to open auto start registry key for reading", exc_info=True)
		return []

	try:
		conf: List[str] = winreg.QueryValueEx(k, "Configuration")[0].split(",")
	except FileNotFoundError:
		log.debug("Unable to find auto start configuration")
	except WindowsError:
		log.error("Unable to query auto start configuration", exc_info=True)
	else:
		if not conf[0]:
			# "".split(",") returns [""], so remove the empty string.
			del conf[0]
		return conf
	return []


def setAutoStart(hkey: _KeyType, enable: bool) -> None:
	"""Raises `Union[WindowsError, FileNotFoundError]`"""
	conf = _getAutoStartConfiguration(hkey)
	currentlyEnabled = APP_KEY_NAME in conf
	if enable and not currentlyEnabled:
		conf.append(APP_KEY_NAME)
		changed = True
	elif not enable:
		try:
			conf.remove(APP_KEY_NAME)
			changed = True
		except ValueError:
			pass
	if changed:
		k = winreg.OpenKey(
			hkey,
			ROOT_KEY,
			0,
			winreg.KEY_READ | winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY
		)
		winreg.SetValueEx(k, "Configuration", None, winreg.REG_SZ,
			",".join(conf))
