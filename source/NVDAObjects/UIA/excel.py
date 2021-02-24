# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2018-2021 NV Access Limited

import UIAHandler
import controlTypes
from . import UIA


class ExcelCell(UIA):

	shouldAllowDuplicateUIAFocusEvent = True

	name = ""
	role = controlTypes.ROLE_TABLECELL
	rowHeaderText = None
	columnHeaderText = None

	def _hasSelection(self):
		return (
			self.selectionContainer
			and 1 < self.selectionContainer.getSelectedItemsCount()
		)

	def _get_value(self):
		if self._hasSelection():
			return
		return super().value

	def _get_description(self):
		if self._hasSelection():
			return
		return self.UIAElement.currentItemStatus

	def _get_cellCoordsText(self):
		if self._hasSelection():
			sc = self._getUIACacheablePropertyValue(
				UIAHandler.UIA_SelectionItemSelectionContainerPropertyId
			).QueryInterface(
				UIAHandler.IUIAutomationElement
			)

			firstSelected = sc.GetCurrentPropertyValue(
				UIAHandler.UIA_Selection2FirstSelectedItemPropertyId
			).QueryInterface(
				UIAHandler.IUIAutomationElement
			)

			firstAddress = firstSelected.GetCurrentPropertyValue(
				UIAHandler.UIA_NamePropertyId
			).replace('"', '')

			firstValue = firstSelected.GetCurrentPropertyValue(
				UIAHandler.UIA_ValueValuePropertyId
			)

			lastSelected = sc.GetCurrentPropertyValue(
				UIAHandler.UIA_Selection2LastSelectedItemPropertyId
			).QueryInterface(
				UIAHandler.IUIAutomationElement
			)

			lastAddress = lastSelected.GetCurrentPropertyValue(
				UIAHandler.UIA_NamePropertyId
			).replace('"', '')

			lastValue = lastSelected.GetCurrentPropertyValue(
				UIAHandler.UIA_ValueValuePropertyId
			)

			return pgettext(
				"excel-UIA",
				# Translators: Excel, report range of cell coordinates
				"{firstAddress} {firstValue} through {lastAddress} {lastValue}"
			).format(
				firstAddress=firstAddress,
				firstValue=firstValue,
				lastAddress=lastAddress,
				lastValue=lastValue
			)
		name = super().name
		# Later builds of Excel 2016 quote the letter coordinate.
		# We don't want the quotes.
		name = name.replace('"', '')
		return name


class ExcelWorksheet(UIA):
	role = controlTypes.ROLE_TABLE

	def _get_name(self):
		return super().parent.name

	def _get_parent(self):
		return super().parent.parent


class CellEdit(UIA):
	name = ""


class BadExcelFormulaEdit(UIA):
	shouldAllowUIAFocusEvent = False
