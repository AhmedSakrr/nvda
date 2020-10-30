# A part of NonVisual Desktop Access (NVDA)
# Copyright (C) 2021 NV Access Limited
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.

import enum


class FillType(enum.IntEnum):
	none = 0
	color = 1
	gradient = 2
	picture = 3
	pattern = 4


FillTypeLabels = {
	# Translators: a style of fill type (to color the inside of a control or text)
	FillType.none: pgettext("UIAHandler.FillType", "none"),
	# Translators: a style of fill type (to color the inside of a control or text)
	FillType.color: pgettext("UIAHandler.FillType", "color"),
	# Translators: a style of fill type (to color the inside of a control or text)
	FillType.gradient: pgettext("UIAHandler.FillType", "gradient"),
	# Translators: a style of fill type (to color the inside of a control or text)
	FillType.picture: pgettext("UIAHandler.FillType", "picture"),
	# Translators: a style of fill type (to color the inside of a control or text)
	FillType.pattern: pgettext("UIAHandler.FillType", "pattern"),
}
