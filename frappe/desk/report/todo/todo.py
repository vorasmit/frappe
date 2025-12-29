# Copyright (c) 2015, Frappe Technologies Pvt. Ltd. and Contributors
# License: MIT. See LICENSE

import frappe
from frappe import _
from frappe.utils import getdate


def execute(filters=None):
	priority_map = {"High": 3, "Medium": 2, "Low": 1}

	todo_list = frappe.get_list(
		"ToDo",
		fields=[
			"name",
			"date",
			"description",
			"priority",
			"reference_type",
			"reference_name",
			"assigned_by",
			"owner",
		],
		filters={"status": "Open"},
	)

	todo_list.sort(
		key=lambda todo: (
			priority_map.get(todo.priority, 0),
			(todo.date and getdate(todo.date)) or getdate("1900-01-01"),
		),
		reverse=True,
	)

	columns = [
		_("ID") + ":Link/ToDo:90",
		_("Priority") + "::60",
		_("Date") + ":Date",
		_("Description") + "::150",
		_("Assigned To/Owner") + ":Data:120",
		_("Assigned By") + ":Data:120",
		_("Reference") + "::200",
	]

	result = []
	for todo in todo_list:
		if todo.owner == frappe.session.user or todo.assigned_by == frappe.session.user:
			if todo.reference_type:
				todo.reference = """<a href="/desk/Form/{}/{}">{}: {}</a>""".format(
					todo.reference_type,
					todo.reference_name,
					todo.reference_type,
					todo.reference_name,
				)
			else:
				todo.reference = None
			result.append(
				[
					todo.name,
					todo.priority,
					todo.date,
					todo.description,
					todo.owner,
					todo.assigned_by,
					todo.reference,
				]
			)

	return columns, result


def get_xlsx_styles(data, xlsx_data, filters, metadata=None):
	"""
	Configure Excel styling for ToDo report export.

	This function demonstrates the declarative Excel styling approach:
	- Highlights High priority tasks in red
	- Highlights Medium priority in orange
	- Shows overdue tasks with a light red background
	- Makes the header row bold with blue background

	Args:
		data: Report data (columns, result)
		xlsx_data: Processed Excel data
		filters: Report filters
		metadata: Optional dict with columns, row_map, total_row_idx, filter_count
	"""
	from datetime import date

	from frappe.utils.xlsx_styles import XLSXStyleBuilder

	builder = XLSXStyleBuilder()

	# Register reusable styles
	builder.register_style(
		"header", font={"bold": True, "size": 12, "color": "FFFFFF"}, fill={"fgColor": "366092", "fill_type": "solid"}
	)

	builder.register_style("high_priority", font={"bold": True, "color": "C00000"})

	builder.register_style("medium_priority", font={"color": "FF6600"})

	builder.register_style("overdue", fill={"fgColor": "FFE6E6", "fill_type": "solid"})

	builder.register_style("total", font={"bold": True}, fill={"fgColor": "D9D9D9", "fill_type": "solid"})

	# Style the header row (first row in xlsx_data if no filters, otherwise after filters)
	# The header_index is passed by the export function, but we can infer it
	# For simplicity, we'll style row 0 which should be the header
	builder.style_row(0, "header")

	# Add conditional styles for priorities
	# Column 1 is Priority column (0-indexed)
	builder.add_conditional_style(lambda r, c, v: c == 1 and v == "High", "high_priority")

	builder.add_conditional_style(lambda r, c, v: c == 1 and v == "Medium", "medium_priority")

	# Highlight overdue tasks (Column 2 is Date column)
	def is_overdue(row_idx, col_idx, value):
		# Only check date column
		if col_idx != 2:
			return False
		# Skip header row
		if row_idx == 0:
			return False
		# Check if date exists and is in the past
		if value and isinstance(value, date):
			return value < date.today()
		return False

	builder.add_conditional_style(is_overdue, "overdue")

	return builder.build()

