#!/usr/bin/env python
# -*- coding: utf8 -*-

import pyodbc
import sys
import wx
from wx import xrc
ctrl = xrc.XRCCTRL

import workdays
import datetime as dt
import General as gn
import Database as db
import Item


class SchedulingTab(object):
	def init_scheduling_tab(self):
		#Bindings
		self.Bind(wx.EVT_CHOICE, self.on_choice_table, id=xrc.XRCID('choice:which_table'))
		
		self.Bind(wx.EVT_BUTTON, self.on_click_proto_request_dates, id=xrc.XRCID('button:proto_request_dates'))
		self.Bind(wx.EVT_BUTTON, self.on_click_calc_and_set_requested_de_release, id=xrc.XRCID('button:calc_and_set_requested_de_release'))



	def on_click_calc_and_set_requested_de_release(self, event):
		records = db.query('''
			SELECT
				id,
				date_basic_start
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release IS NULL AND
				date_basic_start IS NOT NULL AND
				material <> 'SPARTCOLS' AND
				status <> 'Canceled'
			''')

		for record in records:
			id, date_basic_start = record
			
			calc_requested_de_release = workdays.workday(date_basic_start, -23)
			
			db.update_order('orders.target_dates', id, 'requested_de_release', calc_requested_de_release, '{} (Auto)'.format(gn.user))
			
		wx.MessageBox("{} requested_de_release fields have been calculated and set.".format(len(records)), 'Notice', wx.OK | wx.ICON_INFORMATION)


	def on_click_proto_request_dates(self, event):
		records = db.query('''
			SELECT
				id,
				
				CASE
					WHEN bpcs_sales_order IS NULL THEN sales_order
					WHEN sales_order IS NULL THEN bpcs_sales_order
					ELSE sales_order + '/' + bpcs_sales_order
				END AS sales_order,

				CASE
					WHEN bpcs_line_up IS NULL THEN item
					WHEN item IS NULL THEN bpcs_line_up
					ELSE item + '/' + bpcs_line_up
				END AS item,

				CASE
					WHEN bpcs_item IS NULL THEN production_order
					WHEN production_order IS NULL THEN bpcs_item
					ELSE production_order + '/' + bpcs_item
				END AS production_order,

				CASE
					WHEN bpcs_family IS NULL THEN material
					WHEN material IS NULL THEN bpcs_family
					ELSE material + '/' + bpcs_family
				END AS material,

				date_basic_start,
				date_requested_de_release
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release IS NULL AND
				production_order IS NOT NULL AND
				material <> 'SPARTCOLS' AND
				status <> 'Canceled'
			ORDER BY
				date_basic_start ASC
			''')
		
		output = '	'.join(['Sales Order', 'Item', 'Production Order', 'When Got ProdOrd', 'Material', 'Basic Start', 'Ricky Request', 'Calculated Request'])
		output += '\n'
		
		for index, record in enumerate(records):
			id, sales_order, item, production_order, material, basic_start, ricky_request = record
			
			when_got_prod_ord = db.query('''
				SELECT TOP 1
					when_changed
				FROM
					orders.changes
				WHERE
					field = 'production_order' AND
					table_id = {}
				ORDER BY
					id DESC
				'''.format(id))


			
			try:
				calc_request = workdays.workday(basic_start, -23)
			except:
				calc_request = ''



			if when_got_prod_ord:
				when_got_prod_ord = when_got_prod_ord[0].strftime('%m/%d/%Y')
			else:
				when_got_prod_ord = ''

			if basic_start == None:
				basic_start = ''
				
			if ricky_request == None:
				ricky_request = ''
			
			try: basic_start = basic_start.strftime('%m/%d/%Y')
			except: pass

			try: ricky_request = ricky_request.strftime('%m/%d/%Y')
			except: pass

			try: calc_request = calc_request.strftime('%m/%d/%Y')
			except: pass

			output += '{}	'.format(sales_order)
			output += '{}	'.format(item)
			output += '{}	'.format(production_order)
			output += '{}	'.format(when_got_prod_ord)
			output += '{}	'.format(material)
			output += '{}	'.format(basic_start)
			output += '{}	'.format(ricky_request)
			output += '{}\n'.format(calc_request)
			
			
		ctrl(self, 'text:proto_request_dates').SetValue(output)

