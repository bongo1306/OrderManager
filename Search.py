#!/usr/bin/env python
# -*- coding: utf8 -*-

import pyodbc
import sys
import wx
from wx import xrc
ctrl = xrc.XRCCTRL

import datetime as dt
import TweakedGrid
import General as gn
import Database as db


class SearchTab(object):
	def init_search_tab(self):
		#this will hold the fleeting search criteria tables
		self.table_search_criteria = None
		
		#Bindings
		self.Bind(wx.EVT_CHOICE, self.on_choice_table, id=xrc.XRCID('choice:which_table'))
		self.Bind(wx.EVT_CHECKBOX, self.on_choice_table, id=xrc.XRCID('checkbox:display_alphabetically'))
		
		#tables or views the user can search in
		tables = ('orders.root', 'orders.view_systems', 'dbo.orders', 'dbo.view_orders_old')
		
		ctrl(self, 'choice:which_table').AppendItems(tables)
		ctrl(self, 'choice:which_table').SetStringSelection('orders.view_systems')
		self.on_choice_table()
		


	def on_choice_table(self, event=None):
		table_panel = ctrl(self, 'panel:search_criteria')
		
		#remove the table if there is already one there
		if self.table_search_criteria != None:
			self.table_search_criteria.Destroy()
		
		self.table_search_criteria = TweakedGrid.TweakedGrid(table_panel)
		
		table_to_search = ctrl(self, 'choice:which_table').GetStringSelection()

		columns = list(db.get_table_column_names(table_to_search, presentable=False))
		
		#manually specify columns for some tables...
		try:
			columns = list(custom_search_table_columns[table_to_search])
		except:
			pass
		
		for column_index, column in enumerate(columns):
			if '_spacer_' in column:
				column = ''
			columns[column_index] = '{}'.format(column)


		#sort columns alphabetically if user wants
		if ctrl(self, 'checkbox:display_alphabetically').Value == True:
			columns = filter(None, columns)
			columns.sort()

		self.table_search_criteria.CreateGrid(len(columns), 2)
		self.table_search_criteria.SetRowLabelSize(0)
		self.table_search_criteria.SetColLabelValue(0, 'Field')
		self.table_search_criteria.SetColLabelValue(1, 'Criteria')
		
		for column_index, column in enumerate(columns):
			if column != '':
				self.table_search_criteria.SetCellValue(column_index, 0, column)
				self.table_search_criteria.SetCellValue(column_index, 1, '')
			else:
				self.table_search_criteria.SetReadOnly(column_index, 1)
				#self.table_search_criteria.SetCellValue(column_index, 0, column)
				#self.table_search_criteria.SetCellValue(column_index, 1, 'Fields from {}:'.format(column[1:]))
		
		self.table_search_criteria.AutoSize()
		self.table_search_criteria.EnableDragRowSize(False)
		
		self.table_search_criteria.Bind(wx.EVT_SIZE, self.on_size_criteria_table)
		self.table_search_criteria.Bind(wx.grid.EVT_GRID_CELL_CHANGE, self.on_change_grid_cell)
		#self.table_search_criteria.Bind(wx.EVT_CHAR, on_change_grid_cell)
		
		for row in range(len(columns)):
			self.table_search_criteria.SetReadOnly(row, 0)
			#self.table_search_criteria.SetCellAlignment(row, 0, wx.ALIGN_RIGHT, wx.ALIGN_CENTRE)
		
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add(self.table_search_criteria, 1, wx.EXPAND)
		table_panel.SetSizer(sizer)
		
		table_panel.Layout()


	def on_size_criteria_table(self, event):
		table = event.GetEventObject()
		table.SetColSize(1, table.GetSize()[0] - table.GetColSize(0) - wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_X))

		event.Skip()


	def on_change_grid_cell(self, event):
		ctrl(self, 'text:sql_query').SetValue(self.generate_sql_query())
		event.Skip()


	def generate_sql_query(self):
		print 'generating sql query'
		table_to_search = ctrl(self, 'choice:which_table').GetStringSelection()

		sql = "SELECT "
		
		#limit the records pulled if desired
		#if ctrl(General.app.main_frame, 'choice:search_limit').GetStringSelection() != '(no limit)':
		#	sql += "TOP {} ".format(int(ctrl(General.app.main_frame, 'choice:search_limit').GetStringSelection().split(' ')[0]))

		sql_criteria = ''

		#loop through fields
		for row in range(self.table_search_criteria.GetNumberRows()):
			if self.table_search_criteria.GetCellValue(row, 1) != '':
				if self.table_search_criteria.GetCellValue(row, 0) != '': #skip over spacer cell (between tables)
					sql_criteria += self.search_criteria_to_sql(
							column = self.table_search_criteria.GetCellValue(row, 0), 
							criteria = self.table_search_criteria.GetCellValue(row, 1))

		columns = list(db.get_table_column_names(table_to_search, presentable=False))
		fields_to_select = ', '.join(columns)
		
		sql += "{} FROM {} ".format(fields_to_select, table_to_search)
		sql += 'WHERE {}'.format(sql_criteria[:-4])

		##limit the records pulled if desired
		#if ctrl(General.app.main_frame, 'choice:search_limit').GetStringSelection() != '(no limit)':
		#	sql += "LIMIT {}".format(int(ctrl(General.app.main_frame, 'choice:search_limit').GetStringSelection().split(' ')[0]))

		return sql



	def search_criteria_to_sql(self, column, criteria):
		operators = ['<=', '>=', '!=', '<>', '=', '<', '>']
		tokens = [' AND ', ' OR ', '...']
		
		#split string by tokens
		criteria_parts = []
		previous_split_index = 0
		for index in range(len(criteria)):
			for token in tokens:
				if criteria[index:index + len(token)] == token or index == len(criteria)-1:
					if token == '...':
						lower_limit = '>= {}'.format(criteria[previous_split_index:index].rstrip())

						#find next space character to signify end of ... statement
						for char_index in range(index+len(token), len(criteria)-1):
							space_index = char_index+1
							if criteria[char_index] == ' ':
								space_index -= 1
								break

						upper_limit = ' AND <= {} '.format(criteria[index+len(token):space_index+1].rstrip())
						
						criteria_parts.append(lower_limit)
						criteria_parts.append(upper_limit)
						previous_split_index = space_index+1
						break
						
					else:
						criteria_parts.append(criteria[previous_split_index:index+1].rstrip())
						previous_split_index = index
						break
					
		#remove any '' criteria_parts
		criteria_parts = [value for value in criteria_parts if value != '']

		sql_criterias = []
		sql_text = '('

		for criteria_part in criteria_parts:
			#determine and strip out token from criteria
			token_found = None
			for token in tokens:
				if token in criteria_part:
					token_found = token
					criteria_part = criteria_part.replace(token, '')
					break

			#determine and strip out operator from criteria
			operator_found = None
			for operator in operators:
				if operator in criteria_part:
					operator_found = operator
					criteria_part = criteria_part.replace(operator, '').strip()
					break

			#force not equal sign to be ANSI compliant
			try:
				operator_found = operator_found.replace('!=', '<>')
			except:
				pass

			#if not criteria, just '=' sign then make it check if null
			#print 'criteria_part',criteria_part
			if operator_found == '=' and criteria_part == '':
				criteria_part = 'IS NULL'
				operator_found = None

			elif operator_found == '<>' and criteria_part == '':
				criteria_part = 'IS NOT NULL'
				operator_found = None

			else:
				#is it a date?
				criteria_part_is_date = False
				if criteria_part.count('/') > 1:
					criteria_part_is_date = True

				#is it a number?
				criteria_part_is_number = True
				try:
					criteria_part = str(float(criteria_part)).rstrip('.0')
				except:
					if criteria_part_is_date:
						criteria_part = "'{}'".format(criteria_part)
					else:
						criteria_part = "'%{}%'".format(criteria_part)
					criteria_part_is_number = False

				#include all the time in the day if checking <= a date
				if criteria_part_is_date and operator_found == '<=':
					criteria_part = "{} 23:59:59'".format(criteria_part[:-1])

				#if no operators found, it should be a LIKE or = depending if string or number
				if operator_found == None:
					if criteria_part_is_number:
						operator_found = '='
					else:
						if criteria_part_is_date:
							operator_found = '='
						else:
							operator_found = 'LIKE'

			#build up the SQL
			if token_found:
				sql_text += token_found

			if not operator_found:
				operator_found = ''

			sql_text += '{} {} {}'.format(column, operator_found, criteria_part)

		sql_text += ') AND '
		
		return sql_text

