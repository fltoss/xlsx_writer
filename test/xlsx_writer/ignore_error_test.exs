defmodule XlsxWriter.IgnoreErrorTest do
  use ExUnit.Case, async: true

  test "adds an inclusive range without discarding previous instructions" do
    sheet = XlsxWriter.new_sheet("Codes") |> XlsxWriter.write(1, 0, "0101")
    {"Codes", previous} = sheet

    assert {"Codes",
            [
              {:ignore_error_range, 1, 0, 99, 0, :number_stored_as_text}
              | ^previous
            ]} =
             XlsxWriter.ignore_error_range(
               sheet,
               1,
               0,
               99,
               0,
               :number_stored_as_text
             )
  end

  test "rejects invalid row and column indices before calling the NIF" do
    sheet = XlsxWriter.new_sheet("Codes")

    for row <- [-1, 1_048_576, 1.5, "1", nil] do
      assert_raise ArgumentError, ~r/Row index/, fn ->
        XlsxWriter.ignore_error_range(
          sheet,
          row,
          0,
          9,
          0,
          :number_stored_as_text
        )
      end

      assert_raise ArgumentError, ~r/Row index/, fn ->
        XlsxWriter.ignore_error_range(
          sheet,
          0,
          0,
          row,
          0,
          :number_stored_as_text
        )
      end
    end

    for col <- [-1, 16_384, 1.5, "1", nil] do
      assert_raise ArgumentError, ~r/Column index/, fn ->
        XlsxWriter.ignore_error_range(
          sheet,
          0,
          col,
          9,
          0,
          :number_stored_as_text
        )
      end

      assert_raise ArgumentError, ~r/Column index/, fn ->
        XlsxWriter.ignore_error_range(
          sheet,
          0,
          0,
          9,
          col,
          :number_stored_as_text
        )
      end
    end
  end

  test "rejects reversed ranges and unsupported error types" do
    sheet = XlsxWriter.new_sheet("Codes")

    for {r1, c1, r2, c2} <- [{9, 0, 1, 0}, {0, 2, 0, 1}] do
      assert_raise ArgumentError, ~r/Range start/, fn ->
        XlsxWriter.ignore_error_range(
          sheet,
          r1,
          c1,
          r2,
          c2,
          :number_stored_as_text
        )
      end
    end

    assert_raise ArgumentError, ~r/Unsupported error type/, fn ->
      apply(XlsxWriter, :ignore_error_range, [sheet, 0, 0, 0, 0, :all])
    end
  end

  test "accepts the last Excel row and column" do
    assert {"Codes", [_]} =
             XlsxWriter.new_sheet("Codes")
             |> XlsxWriter.ignore_error_range(
               1_048_575,
               16_383,
               1_048_575,
               16_383,
               :number_stored_as_text
             )
  end

  @tag :native_integration
  test "writes scoped warning suppression and preserves text identifiers" do
    sheet =
      XlsxWriter.new_sheet("Codes")
      |> XlsxWriter.write(6, 4, "0101", format: [{:num_format, "@"}])
      |> XlsxWriter.ignore_error_range(6, 4, 99, 4, :number_stored_as_text)

    {:ok, bytes} =
      XlsxWriter.generate([sheet, XlsxWriter.new_sheet("Unchanged")])

    {:ok, entries} = :zip.extract(bytes, [:memory])
    entries = Map.new(entries)
    xml = entries[~c"xl/worksheets/sheet1.xml"]
    assert xml =~ ~s(<ignoredError sqref="E7:E100" numberStoredAsText="1"/>)
    assert xml =~ ~r/<c\b[^>]*r="E7"[^>]*t="s"/
    assert entries[~c"xl/sharedStrings.xml"] =~ "<t>0101</t>"
    refute entries[~c"xl/worksheets/sheet2.xml"] =~ "ignoredErrors"
  end

  @tag :native_integration
  test "supports a single cell and multiple disjoint ranges" do
    sheet =
      XlsxWriter.new_sheet("Codes")
      |> XlsxWriter.ignore_error_range(0, 0, 0, 0, :number_stored_as_text)
      |> XlsxWriter.ignore_error_range(
        6,
        4,
        1_048_575,
        4,
        :number_stored_as_text
      )

    {:ok, bytes} = XlsxWriter.generate([sheet])
    {:ok, entries} = :zip.extract(bytes, [:memory])
    xml = Map.new(entries)[~c"xl/worksheets/sheet1.xml"]
    assert xml =~ ~s(sqref="A1 E7:E1048576")
    assert xml =~ ~s(numberStoredAsText="1")
  end
end
