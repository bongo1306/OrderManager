#!/usr/bin/env python
# -*- coding: utf8 -*-

#Change log:
#v0.7 - 1/8/14
#	*Copying rows now uses \r\n for line breaks to be backwards compatible with other software
#v0.6 - 1/2/14
#	*Added function to clean up sorting symbols from headers
#	*Added function to trigger blank filter control
#v0.5 - 8/21/13
#	*Fixed date sorting
#v0.4 - 8/8/13
#	*Save FileDialog now saves chosen directory as CWD
#v0.3 - 7/17/13
#	*Fixed bug where export would fail if unicode in list. Just converted unicode to ascii.
#v0.2 - 7/15/13
#	*Removed system alert sound when typing into list for Quick Filtering
#	*Added Ctrl+C support (copies selected rows)
#	*Fixed bug when sorting an empty list
#	*Allowed print and export functions to be called passing an event
#v0.1 - 7/12/13
#	*Initial release

import wx
import operator
import sys
import os
import csv
import dateutil.parser as date_parser

import wx.html
from wx.html import HtmlEasyPrinting


class BetterListCtrl(wx.ListCtrl):
	def __init__(self):
		pre = wx.PreListCtrl()
		self.PostCreate(pre)
		self.Bind(wx.EVT_WINDOW_CREATE, self.on_create)
		self.Bind(wx.EVT_CONTEXT_MENU, self.on_context)
		self.Bind(wx.EVT_CHAR, self.on_char)
		self.Bind(wx.EVT_KEY_DOWN, self.on_key)

		self.Bind(wx.EVT_KILL_FOCUS, self.on_focus_change)
		self.Bind(wx.EVT_SET_FOCUS, self.on_focus_change)
		
		self.Bind(wx.EVT_LIST_DELETE_ALL_ITEMS, self.on_items_deleted)
		self.Bind(wx.EVT_LIST_COL_CLICK, self.sort_by_column)
		self.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, self.calculate_right_clicked_cell_value)
		
		self.quick_filter_frame = None
		self.original_list_data = None
		self.ignore_delete_all_items_event = False
		
		self.right_clicked_cell_value = ''
		
		#override these settings if you want
		self.printer_header = ''
		self.printer_font_size = 9
		self.printer_paper_type = wx.PAPER_LETTER	#wx.PAPER_11X17
		self.printer_paper_orientation = wx.LANDSCAPE	#wx.PORTRAIT
		self.printer_paper_margins = (4, 4, 4, 4)	#specify like (0, 0, 0, 0), otherwise None means whatever default is


	def on_create(self, event):
		self.Unbind(wx.EVT_WINDOW_CREATE)
		#extra initialization here


	def clean_headers(self):
		#removes sorting symbols from headers
		headers = []
		for col in range(self.GetColumnCount()):
			headers.append(self.GetColumn(col).GetText())

		if u'↓' in ''.join(headers) or u'↑' in ''.join(headers):
			self.DeleteAllColumns()
			for header_index, header in enumerate(headers):
				self.InsertColumn(header_index, header.replace(u'↓', '').replace(u'↑', '').strip())


	def on_key(self, event):
		#copy selected rows if Ctrl+C is pressed
		if event.ControlDown() and event.GetKeyCode() == 67:
			self.copy_selected_rows()
		
		event.Skip()


	def on_items_deleted(self, event):
		if self.ignore_delete_all_items_event == False:
			#print 'All the items were manually deleted. Closing QF frame and Nulling original_list_data'
			
			if self.quick_filter_frame:
				self.quick_filter_frame.Close()
				
			self.original_list_data = None

	
	def on_focus_change(self, event):
		#show or hide the QF frame depending on if the list is visible (it might be hidden in a tab)
		if self.quick_filter_frame != None:
			if self.IsShownOnScreen(): #maybe also: AND if parent window of this list is IN FOCUS!
				self.quick_filter_frame.Show()
			else:
				self.quick_filter_frame.Hide()
				
		event.Skip()


	def on_char(self, event):
		if event.GetKeyCode() > 32 and event.GetKeyCode() < 127:
			#a character was typed, so create QF frame if it doesn't exist
			if self.quick_filter_frame == None:
				#build a backup of the original list
				self.original_list_data = OriginalListData(self)
				
				self.quick_filter_frame = QuickFilterFrame(self)

				list_size =  self.GetClientSize()
				frame_size = self.quick_filter_frame.GetSize()

				self.quick_filter_frame.Move(self.ClientToScreen((list_size[0]-frame_size[0], list_size[1]-frame_size[1])))
				self.quick_filter_frame.Show(True)
				
			self.quick_filter_frame.text.SetFocus()
			self.quick_filter_frame.text.AppendText(chr(event.GetKeyCode()))
			
			#self.build_filtered_list(self.quick_filter_frame.text.GetValue())
			
		else:
			event.Skip()


	def filter_list(self, event):
		if self.quick_filter_frame == None:
			#build a backup of the original list
			self.original_list_data = OriginalListData(self)
			
			self.quick_filter_frame = QuickFilterFrame(self)

			list_size =  self.GetClientSize()
			frame_size = self.quick_filter_frame.GetSize()

			self.quick_filter_frame.Move(self.ClientToScreen((list_size[0]-frame_size[0], list_size[1]-frame_size[1])))
			self.quick_filter_frame.Show(True)
			
			self.quick_filter_frame.text.SetFocus()
			self.quick_filter_frame.text.SetValue(" ")
			self.quick_filter_frame.text.AppendText("")
			
		else:
			self.quick_filter_frame.Close()

	
	def build_filtered_list(self, filter_string):
		#print 'building filtered list using filter:', filter_string
		
		#print 'figure out how to prevent list repainting until filter complete?'
		self.Hide()
		
		self.ignore_delete_all_items_event = True
		self.DeleteAllItems()
		self.ignore_delete_all_items_event = False
		
		#interpret filter string into filter pieces
		filters = filter_string.upper().split(' ')
		try: filters.remove('')
		except: pass
		
		#try to interpret a date filter
		for filter_index, filter in enumerate(filters):
			try:
				if filter.count('/') == 2:
					month, day, year = filter.split('/')
					if int(year) < 1000:
						year = int(year) + 2000
					
					filters[filter_index] = dt.datetime(int(year), int(month), int(day))
					
			except Exception as e:
				#print e
				pass

		entry_index = -1
		for entry in self.original_list_data.entries:
			#filter out list entries based on user's quick filter keywords
			allow_through = True
			for filter in filters:
				key_word_present = False
				for field in entry.fields:
					try: field = field.upper()
					except: pass
					
					#if field is a string, see if filter is in it
					try:
						if filter in field:
							key_word_present = True
					
					except:
						if filter == field:
							key_word_present = True
				
				if key_word_present == False:
					allow_through = False
					break
			
			if allow_through == False:
				continue
			entry_index += 1

			
			self.InsertStringItem(sys.maxint, '#')
			
			for field_index, field in enumerate(entry.fields):
				self.SetStringItem(entry_index, field_index, field)

			if entry.is_selected:
				self.Select(entry_index)
				self.EnsureVisible(entry_index)
			
			self.SetItemBackgroundColour(entry_index, entry.bg_color)
			
		self.Show()
	

	def on_context(self, event):
		#create and show context menu 

		#only do this part the first time so the events are only bound once 
		if not hasattr(self, "entry1"):
			self.entry1 = wx.NewId()
			self.entry2 = wx.NewId()
			self.entry3 = wx.NewId()
			self.entry4 = wx.NewId()
			self.Bind(wx.EVT_MENU, self.on_popup, id=self.entry1)
			self.Bind(wx.EVT_MENU, self.on_popup, id=self.entry2)
			self.Bind(wx.EVT_MENU, self.on_popup, id=self.entry3)
			self.Bind(wx.EVT_MENU, self.on_popup, id=self.entry4)

		#build the menu
		menu = wx.Menu()
		menu.Append(self.entry1, u'Copy "{}"'.format(self.right_clicked_cell_value))
		menu.Append(self.entry2, 'Copy Selected Rows')
		menu.Append(self.entry3, 'Export List')
		menu.Append(self.entry4, 'Print List')

		#show the popup menu
		self.PopupMenu(menu)
		menu.Destroy()


	def on_popup(self, event):
		itemId = event.GetId()
		menu = event.GetEventObject()
		menuItem = menu.FindItemById(itemId)
		print menuItem.GetLabel()
		
		if 'Copy "' in menuItem.GetLabel():
			self.copy_right_clicked_cell_value()
			
		elif 'Copy Selected Rows' in menuItem.GetLabel():
			self.copy_selected_rows()

		elif 'Export List' in menuItem.GetLabel():
			self.export_list()

		elif 'Print List' in menuItem.GetLabel():
			self.print_list()


	def copy_right_clicked_cell_value(self):
		dataObj = wx.TextDataObject()
		dataObj.SetText(self.right_clicked_cell_value)
		if wx.TheClipboard.Open():
			wx.TheClipboard.SetData(dataObj)
			wx.TheClipboard.Flush()
		else:
			wx.MessageBox("Unable to open the clipboard", "Error")


	def copy_selected_rows(self):
		string_to_copy = u''
		
		#copy list headers
		for col in range(self.GetColumnCount()):
			string_to_copy += u'{}\t'.format(self.GetColumn(col).GetText())
		string_to_copy = string_to_copy.rstrip('\t')
		string_to_copy += '\r\n'
		
		#copy selected rows
		for row in range(self.GetItemCount()):
			if self.IsSelected(row):
				for index, col in enumerate(range(self.GetColumnCount())):
					string_to_copy += u'{}\t'.format(self.GetItem(row, col).GetText())

				string_to_copy = string_to_copy.rstrip('\t')
				string_to_copy += '\r\n'

		string_to_copy = string_to_copy.rstrip('\r\n')
		
		dataObj = wx.TextDataObject()
		dataObj.SetText(string_to_copy)
		if wx.TheClipboard.Open():
			wx.TheClipboard.SetData(dataObj)
			wx.TheClipboard.Flush()
		else:
			wx.MessageBox("Unable to open the clipboard", "Error")


	def export_list(self, event=None):
		#prompt user to choose where to save
		save_dialog = wx.FileDialog(self, message="Export file as ...", 
								defaultDir=os.getcwd(), 
								wildcard="CSV file (*.csv)|*.csv", style=wx.SAVE|wx.OVERWRITE_PROMPT|wx.CHANGE_DIR)

		#show the save dialog and get user's input... if not canceled
		if save_dialog.ShowModal() == wx.ID_OK:
			save_path = save_dialog.GetPath()
			save_dialog.Destroy()
		else:
			save_dialog.Destroy()
			return
			
		with open(save_path, 'wb') as csvfile:
			#writer = csv.writer(csvfile, dialect='excel')
			writer = csv.writer(csvfile)
			
			#write list headers
			fields = []
			for col in range(self.GetColumnCount()):
				value = u'{}'.format(self.GetColumn(col).GetText())
				if col == 0 and value == u'ID':
					value = u'id'
				
				fields.append(value.encode('ascii','replace'))
				
			writer.writerow(fields)
			
			#write rows
			for row in range(self.GetItemCount()):
				
				fields = []
				for index, col in enumerate(range(self.GetColumnCount())):
					value = u'{}'.format(self.GetItem(row, col).GetText())
					fields.append(value.encode('ascii','replace'))

				writer.writerow(fields)
		
		wx.MessageBox('Export completed.', 'Info', wx.OK | wx.ICON_INFORMATION)


	def print_list(self, event=None):
		#html = '''<style type=\"text/css\">td{{font-family:Arial; color:black; font-size:8pt;}}</style>
		#	{}<table border="1" cellspacing="0"><tr>\n'''.format(header)

		html = '''<table border="1" cellspacing="0"><tr>\n'''

		#write out headers
		for index, col in enumerate(range(self.GetColumnCount())):
			html += u'''<th align="left" valign="top">{}</th>\n'''.format(self.GetColumn(col).GetText())
		
		html += '</tr>'
		
		#write out data
		for row in range(self.GetItemCount()):
			html += '<tr>'
			
			for index, col in enumerate(range(self.GetColumnCount())):
				html += '''<td align="left" valign="top" nowrap>{}&nbsp;</td>\n'''.format(self.GetItem(row, col).GetText())
				
			html += '</tr>'
			
		html += '</table>'
		
		printer = HtmlEasyPrinting()
		
		printer.SetHeader('{}, Printed on @DATE@, Page @PAGENUM@ of @PAGESCNT@'.format(self.printer_header).lstrip(', '))
		printer.SetStandardFonts(self.printer_font_size)
		printer.GetPrintData().SetPaperId(self.printer_paper_type)
		printer.GetPrintData().SetOrientation(self.printer_paper_orientation)
		
		if self.printer_paper_margins:
			printer.GetPageSetupData().SetMarginTopLeft((self.printer_paper_margins[0], self.printer_paper_margins[1]))
			printer.GetPageSetupData().SetMarginBottomRight((self.printer_paper_margins[2], self.printer_paper_margins[3]))
			
		printer.PrintText(html)
		

	def calculate_right_clicked_cell_value(self, event):
		spt = wx.GetMousePosition()
		fpt = self.ScreenToClient(spt)
		x, y = fpt
		x += self.GetScrollPos(wx.HORIZONTAL)

		last_col = 0
		for col in range(self.GetColumnCount()) :
			col_width = self.GetColumnWidth(col)

			#Calculate the left and right vertical pixel positions
			# of this current column.
			left_pxl_col = last_col
			right_pxl_col = last_col + col_width - 1

			#Compare mouse click point in control coordinates,
			# (verse screen coordinates) to left and right coordinate of
			# each column consecutively until found.
			if left_pxl_col <= x <= right_pxl_col :
				# Mouse was clicked in the current column "col"; done

				col_selected = col
				break

			col_selected = None

			#prep for next calculation of next column
			last_col = last_col + col_width

		self.right_clicked_cell_value = '{}'.format(self.GetItem(event.GetIndex(), col_selected).GetText())


	def sort_by_column(self, event):
		sort_column_index = event.GetColumn()
		#selected_entry = self.GetFirstSelected()

		headers = []
		column_widths = []
		for col in range(self.GetColumnCount()):
			headers.append(self.GetColumn(col).GetText())
			column_widths.append(self.GetColumnWidth(col))
		
		#if original_list_data exists because of filtering, use it and sort it
		# otherwise just temporarily save the original for sorting and building then forget about it
		if self.original_list_data:
			original_list_data_is_temporary = False
		else:
			original_list_data_is_temporary = True
			self.original_list_data = OriginalListData(self)
			
		#prepare the column we're going to sort by making a copy and turning strings to ints if need be etc.
		# Also put in indexes to sort the real list by
		prepped_column_dict = {}
		for entry_index, entry in enumerate(self.original_list_data.entries):
			field_value = entry.fields[sort_column_index]
			
			#remove decorative characters from field value
			field_value = field_value.replace(u'%', u'').replace(u'$', u'').replace(u',', u'')
			
			#try to convert to datetime object if it's a date
			if '/' in field_value:
				try:
					field_value = str(date_parser.parse(field_value))
				except:
					pass
			
			#try to convert string numbers to legit numbers
			try:
				if len(field_value) == 1:
					try:
						field_value = float(field_value)
					except:
						pass
				
				else:
					try:
						if field_value[0] == '0' and field_value[1] != '.':
							pass #it's not a 'real' number... probaly and old format item number
						else:
							field_value = float(field_value)
					except:
						pass
						
			except:
				pass
			
			prepped_column_dict[entry_index] = field_value

		#↓↑
		
		try:
			if u'↓' in headers[sort_column_index]:
				headers[sort_column_index] = headers[sort_column_index].replace(u'↓', u'↑')
				sorted_prepped_column_indexes = zip(*sorted(prepped_column_dict.iteritems(), key=operator.itemgetter(1), reverse=False))[0]
			
			elif u'↑' in headers[sort_column_index]:
				headers[sort_column_index] = headers[sort_column_index].replace(u'↑', u'↓')
				sorted_prepped_column_indexes = zip(*sorted(prepped_column_dict.iteritems(), key=operator.itemgetter(1), reverse=True))[0]
			
			else:
				headers[sort_column_index] = u'{} {}'.format(headers[sort_column_index], u'↑')
				sorted_prepped_column_indexes = zip(*sorted(prepped_column_dict.iteritems(), key=operator.itemgetter(1), reverse=False))[0]
		except Exception as e:
			sorted_prepped_column_indexes = []
			print e
			
		#sort the original list by keying off the sorted prepped column
		sorted_original_list_data = []
		
		for sorted_index in sorted_prepped_column_indexes:
			sorted_original_list_data.append(self.original_list_data.entries[sorted_index])
		
		self.original_list_data.entries = sorted_original_list_data

		for header_index, header in enumerate(headers):
			if header_index != sort_column_index:
				headers[header_index] = headers[header_index].replace(u'↑', u'').replace(u'↓', u'').strip()
		
		#rebuild list
		self.Hide()
		
		self.DeleteAllColumns()
		for header_index, header in enumerate(headers):
			self.InsertColumn(header_index, header)

		if self.quick_filter_frame:
			self.build_filtered_list(self.quick_filter_frame.text.GetValue())
		else:
			self.build_filtered_list('')
		
		#the build_filtered_list calls list.Show() so hide it again for column width adjusting
		self.Hide()
			
		#return column widths to their previous values
		for column_widths_index, column_width in enumerate(column_widths):
			self.SetColumnWidth(column_widths_index, column_width)
		
		#except for the column we are sorting... as the arrow could make the column not wide enough to display
		self.SetColumnWidth(sort_column_index, wx.LIST_AUTOSIZE_USEHEADER)
		
		self.Show()
		
		#ok, we're done with the temp orginal list so kill it
		if original_list_data_is_temporary == True:
			self.original_list_data = None



class OriginalListData():
	def __init__(self, list_ctrl):
		self.entries = []

		for row in range(list_ctrl.GetItemCount()):
			fields = []
			for index, col in enumerate(range(list_ctrl.GetColumnCount())):
				fields.append(list_ctrl.GetItem(row, col).GetText())
				
			bg_color = list_ctrl.GetItemBackgroundColour(row)
			is_selected = list_ctrl.IsSelected(row)

			self.entries.append(ListEntry(fields, bg_color, is_selected))


class ListEntry():
	def __init__(self, fields, bg_color, is_selected):
		self.fields = fields
		self.bg_color = bg_color
		self.is_selected = is_selected


class QuickFilterFrame(wx.Frame):
	def __init__(self, parent):
		wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size(-1,-1), style = wx.FRAME_FLOAT_ON_PARENT|wx.FRAME_NO_TASKBAR|wx.SIMPLE_BORDER)

		self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
		self.SetFont(wx.Font(10, 70, 90, 90, False, wx.EmptyString))
		
		sizer1 = wx.BoxSizer(wx.HORIZONTAL)
		
		self.panel = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
		sizer2 = wx.BoxSizer(wx.HORIZONTAL)
		
		self.label = wx.StaticText(self.panel, wx.ID_ANY, u"Quick Filter:", wx.DefaultPosition, wx.DefaultSize, 0)
		self.label.Wrap(-1)
		sizer2.Add(self.label, 0, wx.ALIGN_CENTER_VERTICAL|wx.TOP|wx.BOTTOM|wx.LEFT, 4)
		
		self.text = wx.TextCtrl(self.panel, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(150,-1), 0)
		sizer2.Add(self.text, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 4)
		
		self.button = wx.Button(self.panel, wx.ID_ANY, u"Clear", wx.DefaultPosition, wx.DefaultSize, wx.BU_EXACTFIT)
		sizer2.Add(self.button, 0, wx.ALIGN_CENTER_VERTICAL|wx.TOP|wx.BOTTOM|wx.RIGHT, 4)
		
		self.panel.SetSizer(sizer2)
		self.panel.Layout()
		sizer2.Fit(self.panel)
		sizer1.Add(self.panel, 1, wx.EXPAND, 4)
		
		self.SetSizer(sizer1)
		self.Layout()

		self.SetSize(self.GetBestSize())
		
		self.parent = parent
		
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)
		self.text.Bind(wx.EVT_TEXT, self.on_text)
		self.button.Bind(wx.EVT_BUTTON, self.on_click_clear)


	def on_close_frame(self, event):
		#undo filter on list
		self.parent.build_filtered_list('')
		
		self.parent.quick_filter_frame = None
		self.Destroy()


	def on_text(self, event):
		#rebuild list since filter criteria changed
		self.parent.build_filtered_list(self.text.GetValue())
		
		#close QF if they cleared out the text box
		if event.GetEventObject().GetValue() == '':
			self.Close()


	def on_click_clear(self, event):
		self.Close()
