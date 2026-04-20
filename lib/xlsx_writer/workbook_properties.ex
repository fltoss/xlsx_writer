defmodule XlsxWriter.WorkbookProperties do
  @moduledoc """
  Document properties for an Excel workbook.

  These properties are visible in the File > Info section of Excel
  and in the file's metadata.

  ## Fields

  - `:author` - The document author
  - `:title` - The document title
  - `:subject` - The document subject
  - `:manager` - The manager field
  - `:company` - The company name
  - `:category` - The document category
  - `:keywords` - Keywords for the document
  - `:comment` - A comment/description for the document
  - `:status` - The document status
  """

  defstruct author: nil,
            title: nil,
            subject: nil,
            manager: nil,
            company: nil,
            category: nil,
            keywords: nil,
            comment: nil,
            status: nil

  @type t :: %__MODULE__{
          author: String.t() | nil,
          title: String.t() | nil,
          subject: String.t() | nil,
          manager: String.t() | nil,
          company: String.t() | nil,
          category: String.t() | nil,
          keywords: String.t() | nil,
          comment: String.t() | nil,
          status: String.t() | nil
        }
end
