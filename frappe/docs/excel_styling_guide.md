# Excel Export Styling Guide

## Overview

Frappe now supports declarative Excel styling for report exports. This feature allows you to customize the appearance of Excel exports without writing per-cell formatting logic, resulting in clean, performant, and maintainable code.

## Key Features

- **Declarative Approach**: Define styling rules once, apply to many cells
- **Performance Optimized**: Style objects are reused, not recreated for each cell
- **Flexible**: Support for column, row, cell, and conditional styling
- **Backwards Compatible**: Existing exports work without changes
- **Easy to Use**: Simple builder pattern API

## Quick Start

Add a `get_xlsx_styles()` function to your report module:

```python
# my_module/report/my_report/my_report.py

def execute(filters=None):
    # Your existing report logic
    columns = [...]
    data = [...]
    return columns, data


def get_xlsx_styles(data, xlsx_data, filters):
    """Configure Excel styling for this report."""
    from frappe.utils.xlsx_styles import XLSXStyleBuilder

    builder = XLSXStyleBuilder()

    # Define styles
    builder.register_style(
        "header",
        font={"bold": True, "size": 12, "color": "FFFFFF"},
        fill={"fgColor": "366092", "fill_type": "solid"}
    )

    # Apply to header row
    builder.style_row(0, "header")

    return builder.build()
```

## XLSXStyleBuilder API

### Registering Styles

Register a named style for reuse:

```python
builder.register_style(
    "style_name",
    font={"bold": True, "size": 12, "color": "FF0000"},
    fill={"fgColor": "FFFF00", "fill_type": "solid"},
    number_format="#,##0.00",
    alignment={"horizontal": "center", "vertical": "center"},
    border={"left": {...}, "right": {...}}
)
```

**Font Options:**
- `bold`: True/False
- `italic`: True/False
- `underline`: "single", "double", etc.
- `strike`: True/False
- `size`: Font size (e.g., 12)
- `color`: Hex color without # (e.g., "FF0000" for red)
- `name`: Font name (e.g., "Calibri")

**Fill Options:**
- `fgColor`: Foreground hex color (e.g., "FFFF00")
- `fill_type`: "solid", "darkGrid", etc.

**Number Format Examples:**
- `"#,##0.00"`: Thousands separator with 2 decimals
- `"0.00%"`: Percentage with 2 decimals
- `"DD/MM/YYYY"`: Date format
- `"#,##0;[Red](#,##0)"`: Positive/negative with red negatives

**Alignment Options:**
- `horizontal`: "left", "center", "right", "justify"
- `vertical`: "top", "center", "bottom"
- `wrap_text`: True/False

### Applying Styles

#### Column-Wide Styling

```python
# Style entire column 3 (0-indexed)
builder.style_column(3, "currency_style")
```

#### Row-Wide Styling

```python
# Style row 0 (header)
builder.style_row(0, "header_style")

# Style last row (total row)
builder.style_row(-1, "total_style")
```

#### Specific Cell Styling

```python
# Style cell at row 5, column 10
builder.style_cell(5, 10, "highlight_style")
```

#### Conditional Styling

```python
# Highlight negative numbers in red
builder.add_conditional_style(
    lambda row, col, value: isinstance(value, (int, float)) and value < 0,
    "negative_style"
)

# Highlight high priority items
builder.add_conditional_style(
    lambda r, c, v: c == 1 and v == "High",  # Column 1, value is "High"
    "high_priority_style"
)
```

### Method Chaining

All methods return `self` for fluent API:

```python
builder = (
    XLSXStyleBuilder()
    .register_style("s1", font={"bold": True})
    .register_style("s2", font={"italic": True})
    .style_row(0, "s1")
    .style_column(1, "s2")
    .build()
)
```

## Style Precedence

When multiple styles could apply to a cell, they are applied in this order (highest priority first):

1. **Cell Style** - Specific cell styling
2. **Conditional Style** - First matching conditional rule
3. **Row Style** - Row-wide styling
4. **Column Style** - Column-wide styling



## Complete Examples

### Example 1: Financial Report

```python
def get_xlsx_styles(data, xlsx_data, filters):
    from frappe.utils.xlsx_styles import XLSXStyleBuilder

    builder = XLSXStyleBuilder()

    # Define styles
    builder.register_style(
        "header",
        font={"bold": True, "size": 12, "color": "FFFFFF"},
        fill={"fgColor": "366092", "fill_type": "solid"}
    )

    builder.register_style(
        "currency",
        number_format="#,##0.00"
    )

    builder.register_style(
        "negative",
        font={"color": "FF0000"},
        number_format="#,##0.00"
    )

    builder.register_style(
        "total",
        font={"bold": True},
        fill={"fgColor": "D9D9D9", "fill_type": "solid"}
    )

    # Apply styles
    builder.style_row(0, "header")  # Header row
    builder.style_row(-1, "total")  # Last row (totals)

    # Currency formatting for amount columns (assume columns 2-5)
    for col in [2, 3, 4, 5]:
        builder.style_column(col, "currency")

    # Highlight negative values
    builder.add_conditional_style(
        lambda r, c, v: c in [2, 3, 4, 5] and isinstance(v, (int, float)) and v < 0,
        "negative"
    )

    return builder.build()
```

### Example 2: Task/Priority Report

```python
def get_xlsx_styles(data, xlsx_data, filters):
    from datetime import date
    from frappe.utils.xlsx_styles import XLSXStyleBuilder

    builder = XLSXStyleBuilder()

    # Define styles
    builder.register_style(
        "header",
        font={"bold": True, "size": 12, "color": "FFFFFF"},
        fill={"fgColor": "366092", "fill_type": "solid"}
    )

    builder.register_style(
        "high_priority",
        font={"bold": True, "color": "C00000"}
    )

    builder.register_style(
        "medium_priority",
        font={"color": "FF6600"}
    )

    builder.register_style(
        "overdue",
        fill={"fgColor": "FFE6E6", "fill_type": "solid"}
    )

    # Apply styles
    builder.style_row(0, "header")

    # Priority highlighting (assume column 1 is priority)
    builder.add_conditional_style(
        lambda r, c, v: c == 1 and v == "High",
        "high_priority"
    )

    builder.add_conditional_style(
        lambda r, c, v: c == 1 and v == "Medium",
        "medium_priority"
    )

    # Highlight overdue dates (assume column 2 is date)
    def is_overdue(row_idx, col_idx, value):
        if col_idx != 2 or row_idx == 0:  # Skip header
            return False
        if value and isinstance(value, date):
            return value < date.today()
        return False

    builder.add_conditional_style(is_overdue, "overdue")

    return builder.build()
```

### Example 3: Sales Dashboard

```python
def get_xlsx_styles(data, xlsx_data, filters):
    from frappe.utils.xlsx_styles import XLSXStyleBuilder

    builder = XLSXStyleBuilder()

    # Define styles
    builder.register_style(
        "header",
        font={"bold": True, "size": 12, "color": "FFFFFF"},
        fill={"fgColor": "366092", "fill_type": "solid"},
        alignment={"horizontal": "center"}
    )

    builder.register_style(
        "section_header",
        font={"bold": True, "size": 11},
        fill={"fgColor": "E7E6E6", "fill_type": "solid"}
    )

    builder.register_style(
        "percent",
        number_format="0.00%"
    )

    builder.register_style(
        "target_met",
        fill={"fgColor": "C6EFCE", "fill_type": "solid"}
    )

    builder.register_style(
        "target_missed",
        fill={"fgColor": "FFC7CE", "fill_type": "solid"}
    )

    # Apply styles
    builder.style_row(0, "header")

    # Percentage columns
    builder.style_column(4, "percent")
    builder.style_column(5, "percent")

    # Conditional formatting for targets
    builder.add_conditional_style(
        lambda r, c, v: c == 5 and isinstance(v, (int, float)) and v >= 1.0,
        "target_met"
    )

    builder.add_conditional_style(
        lambda r, c, v: c == 5 and isinstance(v, (int, float)) and v < 1.0,
        "target_missed"
    )

    return builder.build()
```

## Parameters Passed to get_xlsx_styles()

Your `get_xlsx_styles()` function receives three parameters:

1. **data** (`frappe._dict`): Report data structure containing:
   - `columns`: Column definitions
   - `result`: Result rows (as dicts)
   - `filters`: Applied filters
   - `message`, `chart`, `report_summary`: Additional report data

2. **xlsx_data** (`list[list]`): The actual data being exported to Excel
   - Includes filter rows (if enabled)
   - Includes header row
   - Includes data rows
   - All HTML is stripped and converted to plain text

3. **filters** (`dict`): The filters applied to the report

## Performance Considerations

### ✅ Good Practices

1. **Register styles once, reuse many times**
   ```python
   # Good
   builder.register_style("currency", number_format="#,##0.00")
   for col in [1, 2, 3, 4]:
       builder.style_column(col, "currency")
   ```

2. **Use column/row styles for bulk operations**
   ```python
   # Good - one style for entire column
   builder.style_column(3, "currency")
   ```

3. **Keep conditional functions simple**
   ```python
   # Good - simple, fast check
   lambda r, c, v: c == 5 and v < 0
   ```

### ❌ Bad Practices

1. **Don't create new style objects in conditions**
   ```python
   # Bad - creates new Font object for each cell
   builder.add_conditional_style(
       lambda r, c, v: v < 0,
       {"font": Font(color="FF0000")}  # Don't do this
   )
   ```

2. **Don't use overly complex conditional functions**
   ```python
   # Bad - expensive operations in condition
   lambda r, c, v: expensive_database_call(v) or complex_calculation(v)
   ```

3. **Don't specify cell styles when column/row styles suffice**
   ```python
   # Bad - styling each cell individually
   for row in range(100):
       for col in range(10):
           builder.style_cell(row, col, "same_style")

   # Good - use column/row styles instead
   for col in range(10):
       builder.style_column(col, "same_style")
   ```

## Troubleshooting

### Styles Not Applied

1. **Check report type**: Only works for standard "Query Report" and "Script Report" types
2. **Check is_standard**: Report must have `is_standard = "Yes"`
3. **Check file location**: Must be in module directory (e.g., `my_module/report/my_report/my_report.py`)
4. **Check error log**: Look for "Excel Style Configuration Error" in Error Log

### Conditional Styles Not Working

1. **Check condition function**: Should return True/False
2. **Check row/column indices**: Remember they are 0-indexed
3. **Check value types**: Use `isinstance()` to verify types
4. **Add error handling**: Wrap complex logic in try/except

### Performance Issues

1. **Simplify conditional functions**: Avoid expensive operations
2. **Use column/row styles**: Instead of many cell styles
3. **Limit conditional rules**: Each rule checks every cell
4. **Profile your code**: Use cProfile to identify bottlenecks

## Migration from PR #34707

If you were using the previous hook-based approach:

**Old Approach:**
```python
def get_xlsx_cell_style(cell_value, column, row, filters, is_total_row):
    # Per-cell function call
    if cell_value < 0:
        return {"font": {"color": "FF0000"}}
```

**New Approach:**
```python
def get_xlsx_styles(data, xlsx_data, filters):
    builder = XLSXStyleBuilder()
    builder.register_style("negative", font={"color": "FF0000"})
    builder.add_conditional_style(
        lambda r, c, v: isinstance(v, (int, float)) and v < 0,
        "negative"
    )
    return builder.build()
```

**Benefits:**
- 10-100x faster (style objects reused, not recreated)
- Cleaner code (declarative vs imperative)
- More flexible (can style headers, filters, totals)
- Easier to test and maintain

## Contributing

Found a bug or want to add features? Contributions welcome!

- Report issues: GitHub Issues
- Submit PRs: Follow Frappe contribution guidelines
- Add tests: See `frappe/tests/test_xlsx_styles.py`

## See Also

- [openpyxl Documentation](https://openpyxl.readthedocs.io/)
- [Frappe Report Documentation](https://frappeframework.com/docs/user/en/desk/reports)
- [Excel Number Formats Reference](https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68)
