#!/usr/bin/env elixir
#
# Comprehensive XlsxWriter Feature Demo
#
# This script demonstrates all available features in XlsxWriter.
# Run with: mix run examples/comprehensive_demo.exs
#

IO.puts("Generating comprehensive XlsxWriter demo...")

# Sheet 1: Data Types
data_types_sheet =
  XlsxWriter.new_sheet("Data Types")
  |> XlsxWriter.set_tab_color("#00B050")
  |> XlsxWriter.write(0, 0, "Data Type Examples", format: [:bold, {:font_size, 14}])
  |> XlsxWriter.write(2, 0, "Type", format: [:bold])
  |> XlsxWriter.write(2, 1, "Example", format: [:bold])
  |> XlsxWriter.write(2, 2, "Notes", format: [:bold])
  # String
  |> XlsxWriter.write(3, 0, "String")
  |> XlsxWriter.write(3, 1, "Hello World")
  |> XlsxWriter.write(3, 2, "Basic text")
  # Integer
  |> XlsxWriter.write(4, 0, "Integer")
  |> XlsxWriter.write(4, 1, 42)
  |> XlsxWriter.write(4, 2, "Whole numbers")
  # Float
  |> XlsxWriter.write(5, 0, "Float")
  |> XlsxWriter.write(5, 1, 3.14159)
  |> XlsxWriter.write(5, 2, "Decimal numbers")
  # Boolean
  |> XlsxWriter.write(6, 0, "Boolean")
  |> XlsxWriter.write_boolean(6, 1, true)
  |> XlsxWriter.write(6, 2, "TRUE/FALSE values")
  # Date
  |> XlsxWriter.write(7, 0, "Date")
  |> XlsxWriter.write(7, 1, ~D[2025-01-15])
  |> XlsxWriter.write(7, 2, "Date values")
  # DateTime
  |> XlsxWriter.write(8, 0, "DateTime")
  |> XlsxWriter.write(8, 1, ~U[2025-01-15 14:30:00Z])
  |> XlsxWriter.write(8, 2, "Date with time")
  # Formula
  |> XlsxWriter.write(9, 0, "Formula")
  |> XlsxWriter.write_formula(9, 1, "=B4+B5")
  |> XlsxWriter.write(9, 2, "Excel formulas")
  # URL
  |> XlsxWriter.write(10, 0, "URL")
  |> XlsxWriter.write_url(10, 1, "https://elixir-lang.org", text: "Elixir Website")
  |> XlsxWriter.write(10, 2, "Clickable links")
  # Blank (formatted)
  |> XlsxWriter.write(11, 0, "Blank")
  |> XlsxWriter.write_blank(11, 1, format: [{:bg_color, "#FFFF00"}])
  |> XlsxWriter.write(11, 2, "Formatted empty cell")
  |> XlsxWriter.autofit()

# Sheet 2: Font Formatting
font_sheet =
  XlsxWriter.new_sheet("Font Formatting")
  |> XlsxWriter.set_tab_color("#0000FF")
  |> XlsxWriter.write(0, 0, "Font Formatting Examples", format: [:bold, {:font_size, 14}])
  # Font styles
  |> XlsxWriter.write(2, 0, "Bold Text", format: [:bold])
  |> XlsxWriter.write(3, 0, "Italic Text", format: [:italic])
  |> XlsxWriter.write(4, 0, "Strikethrough", format: [:strikethrough])
  |> XlsxWriter.write(5, 0, "Underlined", format: [{:underline, :single}])
  |> XlsxWriter.write(6, 0, "Double Underline", format: [{:underline, :double}])
  # Font colors
  |> XlsxWriter.write(8, 0, "Red Text", format: [{:font_color, "#FF0000"}])
  |> XlsxWriter.write(9, 0, "Blue Text", format: [{:font_color, "#0000FF"}])
  |> XlsxWriter.write(10, 0, "Green Text", format: [{:font_color, "#00FF00"}])
  # Font sizes
  |> XlsxWriter.write(12, 0, "Small (10pt)", format: [{:font_size, 10}])
  |> XlsxWriter.write(13, 0, "Medium (14pt)", format: [{:font_size, 14}])
  |> XlsxWriter.write(14, 0, "Large (18pt)", format: [{:font_size, 18}])
  |> XlsxWriter.write(15, 0, "Extra Large (24pt)", format: [{:font_size, 24}])
  # Font families
  |> XlsxWriter.write(17, 0, "Arial", format: [{:font_name, "Arial"}])
  |> XlsxWriter.write(18, 0, "Courier New", format: [{:font_name, "Courier New"}])
  |> XlsxWriter.write(19, 0, "Times New Roman", format: [{:font_name, "Times New Roman"}])
  # Text positioning
  |> XlsxWriter.write(21, 0, "E=mc²", format: [:superscript])
  |> XlsxWriter.write(22, 0, "H₂O", format: [:subscript])
  # Combined formatting
  |> XlsxWriter.write(24, 0, "Bold Red Large", format: [:bold, {:font_color, "#FF0000"}, {:font_size, 16}])
  |> XlsxWriter.set_column_width(0, 30)

# Sheet 3: Cell Borders
borders_sheet =
  XlsxWriter.new_sheet("Borders")
  |> XlsxWriter.set_tab_color("#FF0000")
  |> XlsxWriter.write(0, 0, "Border Examples", format: [:bold, {:font_size, 14}])
  # All sides borders
  |> XlsxWriter.write(2, 0, "Thin Border", format: [{:border, :thin}])
  |> XlsxWriter.write(3, 0, "Medium Border", format: [{:border, :medium}])
  |> XlsxWriter.write(4, 0, "Thick Border", format: [{:border, :thick}])
  |> XlsxWriter.write(5, 0, "Dashed Border", format: [{:border, :dashed}])
  |> XlsxWriter.write(6, 0, "Dotted Border", format: [{:border, :dotted}])
  |> XlsxWriter.write(7, 0, "Double Border", format: [{:border, :double}])
  # Individual sides
  |> XlsxWriter.write(2, 2, "Top Only", format: [{:border_top, :thick}])
  |> XlsxWriter.write(3, 2, "Bottom Only", format: [{:border_bottom, :thick}])
  |> XlsxWriter.write(4, 2, "Left Only", format: [{:border_left, :thick}])
  |> XlsxWriter.write(5, 2, "Right Only", format: [{:border_right, :thick}])
  # Colored borders
  |> XlsxWriter.write(9, 0, "Red Border", format: [{:border, :medium}, {:border_color, "#FF0000"}])
  |> XlsxWriter.write(10, 0, "Blue Border", format: [{:border, :medium}, {:border_color, "#0000FF"}])
  # Multi-colored borders
  |> XlsxWriter.write(12, 0, "Rainbow", format: [
    {:border_top, :medium}, {:border_top_color, "#FF0000"},
    {:border_right, :medium}, {:border_right_color, "#00FF00"},
    {:border_bottom, :medium}, {:border_bottom_color, "#0000FF"},
    {:border_left, :medium}, {:border_left_color, "#FFFF00"}
  ])
  |> XlsxWriter.set_column_range_width(0, 2, 20)

# Sheet 4: Background Colors & Patterns
background_sheet =
  XlsxWriter.new_sheet("Backgrounds")
  |> XlsxWriter.set_tab_color("#800080")
  |> XlsxWriter.write(0, 0, "Background Examples", format: [:bold, {:font_size, 14}])
  # Solid colors
  |> XlsxWriter.write(2, 0, "Red Background", format: [{:bg_color, "#FF0000"}, {:font_color, "#FFFFFF"}])
  |> XlsxWriter.write(3, 0, "Green Background", format: [{:bg_color, "#00FF00"}])
  |> XlsxWriter.write(4, 0, "Blue Background", format: [{:bg_color, "#0000FF"}, {:font_color, "#FFFFFF"}])
  |> XlsxWriter.write(5, 0, "Yellow Background", format: [{:bg_color, "#FFFF00"}])
  |> XlsxWriter.write(6, 0, "Orange Background", format: [{:bg_color, "#FFA500"}])
  |> XlsxWriter.write(7, 0, "Purple Background", format: [{:bg_color, "#800080"}, {:font_color, "#FFFFFF"}])
  # Patterns
  |> XlsxWriter.write(9, 0, "Solid Pattern", format: [{:bg_color, "#CCCCCC"}, {:pattern, :solid}])
  |> XlsxWriter.write(10, 0, "Gray 12.5%", format: [{:bg_color, "#CCCCCC"}, {:pattern, :gray125}])
  |> XlsxWriter.write(11, 0, "Gray 6.25%", format: [{:bg_color, "#CCCCCC"}, {:pattern, :gray0625}])
  |> XlsxWriter.set_column_width(0, 25)

# Sheet 5: Alignment, Text Control & Number Formats
formatting_sheet =
  XlsxWriter.new_sheet("Alignment & Numbers")
  |> XlsxWriter.set_tab_color("#FF6600")
  |> XlsxWriter.write(0, 0, "Alignment, Text Control & Number Formatting", format: [:bold, {:font_size, 14}])
  # Horizontal alignment
  |> XlsxWriter.write(2, 0, "Horizontal:", format: [:bold])
  |> XlsxWriter.write(3, 0, "Left Aligned", format: [{:align, :left}])
  |> XlsxWriter.write(4, 0, "Center Aligned", format: [{:align, :center}])
  |> XlsxWriter.write(5, 0, "Right Aligned", format: [{:align, :right}])
  # Vertical alignment
  |> XlsxWriter.write(2, 1, "Vertical:", format: [:bold])
  |> XlsxWriter.write(3, 1, "Top", format: [{:valign, :top}, {:border, :thin}])
  |> XlsxWriter.write(4, 1, "Center", format: [{:valign, :center}, {:border, :thin}])
  |> XlsxWriter.write(5, 1, "Bottom", format: [{:valign, :bottom}, {:border, :thin}])
  |> XlsxWriter.set_row_range_height(3, 5, 35)
  # Combined alignment
  |> XlsxWriter.write(2, 2, "Combined:", format: [:bold])
  |> XlsxWriter.write(3, 2, "Center + Middle",
    format: [{:align, :center}, {:valign, :center}, {:border, :thin}])
  |> XlsxWriter.write(4, 2, "Right + Bottom",
    format: [{:align, :right}, {:valign, :bottom}, {:border, :thin}])
  |> XlsxWriter.write(5, 2, "Left + Top",
    format: [{:align, :left}, {:valign, :top}, {:border, :thin}])
  # Text wrapping
  |> XlsxWriter.write(7, 0, "Text Wrapping:", format: [:bold])
  |> XlsxWriter.write(8, 0, "This is a long text that will automatically wrap within the cell boundaries",
    format: [:text_wrap])
  |> XlsxWriter.write(8, 1, "Wrapped + Bold + Centered",
    format: [:text_wrap, :bold, {:align, :center}, {:valign, :center}])
  |> XlsxWriter.set_row_height(8, 50)
  # Text rotation
  |> XlsxWriter.write(7, 2, "Rotation:", format: [:bold])
  |> XlsxWriter.write(8, 2, "45 degrees", format: [{:rotation, 45}])
  |> XlsxWriter.write(8, 3, "-45 degrees", format: [{:rotation, -45}])
  |> XlsxWriter.write(8, 4, "Vertical", format: [{:rotation, 270}])
  # Shrink to fit
  |> XlsxWriter.write(10, 0, "Shrink to Fit:", format: [:bold])
  |> XlsxWriter.write(11, 0, "This very long text will shrink to fit the column width automatically",
    format: [:shrink])
  # Text indent
  |> XlsxWriter.write(10, 1, "Text Indent:", format: [:bold])
  |> XlsxWriter.write(11, 1, "No indent")
  |> XlsxWriter.write(12, 1, "Indent 1", format: [{:indent, 1}])
  |> XlsxWriter.write(13, 1, "Indent 2", format: [{:indent, 2}])
  |> XlsxWriter.write(14, 1, "Indent 3", format: [{:indent, 3}])
  # Number formats
  |> XlsxWriter.write(16, 0, "Number Formats:", format: [:bold])
  |> XlsxWriter.write(17, 0, "Currency:")
  |> XlsxWriter.write(17, 1, 1234.56, format: [{:num_format, "$#,##0.00"}])
  |> XlsxWriter.write(18, 0, "Percentage:")
  |> XlsxWriter.write(18, 1, 0.75, format: [{:num_format, "0.00%"}])
  |> XlsxWriter.write(19, 0, "Thousands:")
  |> XlsxWriter.write(19, 1, 1234567, format: [{:num_format, "#,##0"}])
  |> XlsxWriter.write(20, 0, "Decimal:")
  |> XlsxWriter.write(20, 1, 3.14159, format: [{:num_format, "0.00"}])
  |> XlsxWriter.write(21, 0, "Scientific:")
  |> XlsxWriter.write(21, 1, 1234.56, format: [{:num_format, "0.00E+00"}])
  |> XlsxWriter.set_column_width(0, 20)
  |> XlsxWriter.set_column_width(1, 20)
  |> XlsxWriter.set_column_width(2, 15)
  |> XlsxWriter.set_column_width(3, 15)
  |> XlsxWriter.set_column_width(4, 10)

# Sheet 6: Layout Features
layout_sheet =
  XlsxWriter.new_sheet("Layout Features")
  |> XlsxWriter.set_tab_color("#4472C4")
  |> XlsxWriter.write(0, 0, "Layout Feature Examples", format: [:bold, {:font_size, 14}])
  # Freeze panes (freeze first two rows)
  |> XlsxWriter.write(1, 0, "Col A", format: [:bold, {:bg_color, "#4472C4"}, {:font_color, "#FFFFFF"}])
  |> XlsxWriter.write(1, 1, "Col B", format: [:bold, {:bg_color, "#4472C4"}, {:font_color, "#FFFFFF"}])
  |> XlsxWriter.write(1, 2, "Col C", format: [:bold, {:bg_color, "#4472C4"}, {:font_color, "#FFFFFF"}])
  |> XlsxWriter.write(1, 3, "Col D", format: [:bold, {:bg_color, "#4472C4"}, {:font_color, "#FFFFFF"}])
  |> XlsxWriter.freeze_panes(2, 0)
  # Autofilter
  |> XlsxWriter.set_autofilter(1, 0, 1, 3)
  # Data rows
  |> XlsxWriter.write(2, 0, "Data 1")
  |> XlsxWriter.write(2, 1, 100)
  |> XlsxWriter.write(2, 2, 200)
  |> XlsxWriter.write(2, 3, 300)
  |> XlsxWriter.write(3, 0, "Data 2")
  |> XlsxWriter.write(3, 1, 150)
  |> XlsxWriter.write(3, 2, 250)
  |> XlsxWriter.write(3, 3, 350)
  # Hidden row
  |> XlsxWriter.write(4, 0, "This row is hidden")
  |> XlsxWriter.hide_row(4)
  |> XlsxWriter.write(5, 0, "Data 4")
  |> XlsxWriter.write(5, 1, 175)
  |> XlsxWriter.write(5, 2, 275)
  |> XlsxWriter.write(5, 3, 375)
  # Hidden column (column E)
  |> XlsxWriter.write(1, 4, "Hidden", format: [:bold])
  |> XlsxWriter.write(2, 4, "Secret")
  |> XlsxWriter.hide_column(4)
  # Column/row range operations
  |> XlsxWriter.set_column_range_width(0, 3, 120)
  |> XlsxWriter.set_row_range_height(2, 5, 25)

# Sheet 7: Merged Cells
merged_sheet =
  XlsxWriter.new_sheet("Merged Cells")
  |> XlsxWriter.set_tab_color("#FFD700")
  # Title spanning columns A-D
  |> XlsxWriter.merge_range(0, 0, 0, 3, "Quarterly Sales Report",
    format: [:bold, {:font_size, 16}, {:align, :center}, {:bg_color, "#4472C4"}, {:font_color, "#FFFFFF"}])
  # Headers
  |> XlsxWriter.write(1, 0, "Product", format: [:bold])
  |> XlsxWriter.write(1, 1, "Q1", format: [:bold])
  |> XlsxWriter.write(1, 2, "Q2", format: [:bold])
  |> XlsxWriter.write(1, 3, "Q3", format: [:bold])
  # Data with vertical merge
  |> XlsxWriter.merge_range(2, 0, 4, 0, "Category A")
  |> XlsxWriter.write(2, 1, 1000)
  |> XlsxWriter.write(2, 2, 1100)
  |> XlsxWriter.write(2, 3, 1200)
  |> XlsxWriter.write(3, 1, 950)
  |> XlsxWriter.write(3, 2, 1050)
  |> XlsxWriter.write(3, 3, 1150)
  |> XlsxWriter.write(4, 1, 1100)
  |> XlsxWriter.write(4, 2, 1200)
  |> XlsxWriter.write(4, 3, 1300)
  # Merge with number and formatting
  |> XlsxWriter.merge_range(6, 1, 6, 3, 12345.67,
    format: [:bold, {:num_format, "$#,##0.00"}, {:align, :center}, {:bg_color, "#FFE699"}])
  |> XlsxWriter.write(6, 0, "Total:")
  |> XlsxWriter.set_column_range_width(0, 3, 120)

# Sheet 8: Complex Example - Invoice
invoice_sheet =
  XlsxWriter.new_sheet("Invoice Example")
  |> XlsxWriter.set_tab_color("#333333")
  # Company header
  |> XlsxWriter.merge_range(0, 0, 1, 3, "ACME Corporation",
    format: [:bold, {:font_size, 20}, {:align, :center}, {:bg_color, "#4472C4"}, {:font_color, "#FFFFFF"}])
  # Invoice details
  |> XlsxWriter.write(3, 0, "Invoice #:", format: [:bold])
  |> XlsxWriter.write(3, 1, "INV-2025-001")
  |> XlsxWriter.write(4, 0, "Date:", format: [:bold])
  |> XlsxWriter.write(4, 1, ~D[2025-01-15])
  # Items table header
  |> XlsxWriter.write(6, 0, "Item", format: [:bold, {:border, :thin}, {:bg_color, "#D9D9D9"}])
  |> XlsxWriter.write(6, 1, "Quantity", format: [:bold, {:border, :thin}, {:bg_color, "#D9D9D9"}])
  |> XlsxWriter.write(6, 2, "Price", format: [:bold, {:border, :thin}, {:bg_color, "#D9D9D9"}])
  |> XlsxWriter.write(6, 3, "Total", format: [:bold, {:border, :thin}, {:bg_color, "#D9D9D9"}])
  # Items
  |> XlsxWriter.write(7, 0, "Widget A", format: [{:border, :thin}])
  |> XlsxWriter.write(7, 1, 10, format: [{:border, :thin}])
  |> XlsxWriter.write(7, 2, 9.99, format: [{:border, :thin}, {:num_format, "$#,##0.00"}])
  |> XlsxWriter.write_formula(7, 3, "=B8*C8")
  |> XlsxWriter.write(8, 0, "Gadget B", format: [{:border, :thin}])
  |> XlsxWriter.write(8, 1, 5, format: [{:border, :thin}])
  |> XlsxWriter.write(8, 2, 24.99, format: [{:border, :thin}, {:num_format, "$#,##0.00"}])
  |> XlsxWriter.write_formula(8, 3, "=B9*C9")
  # Total
  |> XlsxWriter.write(10, 2, "Total:", format: [:bold, {:align, :right}])
  |> XlsxWriter.write_formula(10, 3, "=SUM(D8:D9)")
  # Notes
  |> XlsxWriter.write(12, 0, "Thank you for your business!", format: [:italic, {:font_color, "#666666"}])
  |> XlsxWriter.set_column_width(0, 25)
  |> XlsxWriter.set_column_width(1, 12)
  |> XlsxWriter.set_column_width(2, 15)
  |> XlsxWriter.set_column_width(3, 15)
  |> XlsxWriter.set_row_height(0, 40)

# Sheet 9: Rich Text Formatting
rich_text_sheet =
  XlsxWriter.new_sheet("Rich Text")
  |> XlsxWriter.set_tab_color("#00BFFF")
  |> XlsxWriter.write(0, 0, "Rich Text Examples", format: [:bold, {:font_size, 14}])
  # Bold and normal text
  |> XlsxWriter.write(2, 0, "Mixed Styles:")
  |> XlsxWriter.write_rich_string(2, 1, [
    {"Bold ", [:bold]},
    {"Normal ", []},
    {"Italic", [:italic]}
  ])
  # Colored text segments
  |> XlsxWriter.write(3, 0, "Colors:")
  |> XlsxWriter.write_rich_string(3, 1, [
    {"Red ", [{:font_color, "#FF0000"}]},
    {"Green ", [{:font_color, "#00FF00"}]},
    {"Blue", [{:font_color, "#0000FF"}]}
  ])
  # Scientific notation
  |> XlsxWriter.write(4, 0, "Superscript:")
  |> XlsxWriter.write_rich_string(4, 1, [
    {"E=mc", []},
    {"2", [:superscript]}
  ])
  # Chemical formulas
  |> XlsxWriter.write(5, 0, "Subscript:")
  |> XlsxWriter.write_rich_string(5, 1, [
    {"H", []},
    {"2", [:subscript]},
    {"O", []}
  ])
  # Complex formatting
  |> XlsxWriter.write(6, 0, "Complex:")
  |> XlsxWriter.write_rich_string(6, 1, [
    {"Bold Red ", [:bold, {:font_color, "#FF0000"}]},
    {"Large ", [{:font_size, 16}]},
    {"Underlined", [{:underline, :single}]}
  ])
  # With cell-level formatting
  |> XlsxWriter.write(7, 0, "Cell format:")
  |> XlsxWriter.write_rich_string(7, 1, [
    {"Important: ", [:bold]},
    {"Read carefully", [:italic]}
  ], format: [{:align, :center}, {:bg_color, "#FFFF00"}])
  # Font sizes
  |> XlsxWriter.write(8, 0, "Font sizes:")
  |> XlsxWriter.write_rich_string(8, 1, [
    {"Small ", [{:font_size, 10}]},
    {"Medium ", [{:font_size, 14}]},
    {"Large", [{:font_size, 18}]}
  ])
  # Strikethrough example
  |> XlsxWriter.write(9, 0, "Strikethrough:")
  |> XlsxWriter.write_rich_string(9, 1, [
    {"Original Price: ", []},
    {"$99.99", [:strikethrough]},
    {" Now: $49.99", [:bold, {:font_color, "#FF0000"}]}
  ])
  |> XlsxWriter.autofit()

# Generate the workbook
# Sheet 10: Comments/Notes
comments_sheet =
  XlsxWriter.new_sheet("Comments")
  |> XlsxWriter.set_tab_color("#C00000")
  |> XlsxWriter.write(0, 0, "Cell Comments Demo", format: [:bold, {:font_size, 14}])
  |> XlsxWriter.write_comment(0, 0, "Comments provide additional context and documentation")
  |> XlsxWriter.write(2, 0, "Feature", format: [:bold])
  |> XlsxWriter.write(2, 1, "Example", format: [:bold])
  # Simple comment
  |> XlsxWriter.write(3, 0, "Simple comment")
  |> XlsxWriter.write(3, 1, "Hover to see note")
  |> XlsxWriter.write_comment(3, 1, "This is a basic comment that appears on hover")
  # Comment with author
  |> XlsxWriter.write(4, 0, "With author")
  |> XlsxWriter.write(4, 1, "Review this")
  |> XlsxWriter.write_comment(4, 1, "Please review by end of week", author: "Manager")
  # Visible comment
  |> XlsxWriter.write(5, 0, "Always visible")
  |> XlsxWriter.write(5, 1, "Important!")
  |> XlsxWriter.write_comment(5, 1, "This comment is always shown (not just on hover)",
    visible: true,
    width: 250,
    height: 100
  )
  # Multi-line comment
  |> XlsxWriter.write(6, 0, "Multi-line")
  |> XlsxWriter.write(6, 1, "Instructions")
  |> XlsxWriter.write_comment(6, 1,
    "Line 1: First instruction\nLine 2: Second instruction\nLine 3: Third instruction"
  )
  # Comment with formatting
  |> XlsxWriter.write(7, 0, "With formatting")
  |> XlsxWriter.write(7, 1, 42, format: [:bold, {:bg_color, "#FFFF00"}])
  |> XlsxWriter.write_comment(7, 1, "Comments work with any cell formatting", author: "Developer")
  |> XlsxWriter.set_column_width(0, 20)
  |> XlsxWriter.set_column_width(1, 25)

# Workbook properties (document metadata)
props = %XlsxWriter.WorkbookProperties{
  author: "XlsxWriter Demo",
  title: "Comprehensive Feature Demo",
  subject: "XlsxWriter Library Features",
  company: "Elixir Community",
  category: "Demo",
  keywords: "xlsx, elixir, demo, spreadsheet",
  comment: "Generated by XlsxWriter comprehensive demo script"
}

{:ok, content} = XlsxWriter.generate([
  data_types_sheet,
  font_sheet,
  borders_sheet,
  background_sheet,
  formatting_sheet,
  layout_sheet,
  merged_sheet,
  invoice_sheet,
  rich_text_sheet,
  comments_sheet
], properties: props)

# Write to file in examples folder
output_file = Path.join(__DIR__, "output/comprehensive_demo.xlsx")
File.write!(output_file, content)

IO.puts("✓ Generated #{output_file}")
IO.puts("")
IO.puts("This demo includes:")
IO.puts("  • All data types (strings, numbers, dates, booleans, formulas, URLs)")
IO.puts("  • Font formatting (colors, sizes, styles, families)")
IO.puts("  • Cell borders (all 13 styles, colored, per-side)")
IO.puts("  • Background colors and patterns")
IO.puts("  • Text alignment (horizontal + vertical)")
IO.puts("  • Text wrapping, rotation, shrink to fit, indent")
IO.puts("  • Number formats (currency, percentage, scientific)")
IO.puts("  • Layout features (freeze panes, autofilter, hidden rows/columns)")
IO.puts("  • Column autofit (auto-adjust widths)")
IO.puts("  • Range operations for bulk sizing")
IO.puts("  • Merged cells")
IO.puts("  • Rich text formatting (mixed styles in single cells)")
IO.puts("  • Cell comments/notes")
IO.puts("  • Worksheet tab colors")
IO.puts("  • Workbook properties (author, title, subject, etc.)")
IO.puts("  • A complete invoice example")
