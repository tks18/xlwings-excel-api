---
name: fx_LoadPDFDynamic
category: Table Loaders
tags: [pdf, dynamic, import, columns, table]
description: "Load all pages of a PDF file into a single flattened table by expanding a configurable number of dynamic columns otherwise 100 and removing empty ones."
version: "v2.1"
---
let
    fx_LoadPDFDynamic =
        (pdfFile as text, optional maxColumns as number) as table =>
        let
            // Default value for max columns if not provided
            maxCols = if maxColumns <> null then maxColumns else 100,

            // Helper function to remove columns where all values are null
            Int_fx_TableRemoveNullCols = (sourceTable as table) as table =>
            let
                profile = Table.FromRecords(Table.ToRecords(Table.Profile(sourceTable))),
                nonEmpty = Table.SelectRows(profile, each [NullCount] <> [Count])[Column],
                result = Table.SelectColumns(sourceTable, nonEmpty)
            in
                result,

            // Load PDF as binary
            bin_contents = File.Contents(pdfFile),

            // Extract tables from PDF
            Source = Pdf.Tables(bin_contents),

            // Keep only page-level tables
            FilterPages = Table.SelectRows(Source, each ([Kind] = "Page")),

            // Generate dynamic column names based on maxCols
            pdfCols = List.Transform(List.Numbers(1, maxCols), each "Column" & Text.From(_)),
            GenerateCols = List.Transform(List.Numbers(1, maxCols), each "Col_" & Text.From(_)),

            // Expand and flatten page data
            PdfDumpCols = Table.ExpandTableColumn(FilterPages, "Data", pdfCols, GenerateCols),

            // Drop irrelevant metadata columns
            RemoveUnwantedCols = Table.RemoveColumns(PdfDumpCols, {"Id", "Kind"}),

            // Clean columns with only nulls
            CleanedCols = Int_fx_TableRemoveNullCols(RemoveUnwantedCols)
        in
            CleanedCols,

    // Function type definition with metadata
    fx_type = type function (
        pdfFile as (type text meta [
            Documentation.FieldCaption = "PDF File Path",
            Documentation.FieldDescription = "Path to the PDF file from which to load tables."
        ]),
        optional maxColumns as (type number meta [
            Documentation.FieldCaption = "Max Columns",
            Documentation.FieldDescription = "Maximum number of columns to expand (default = 100)."
        ])
    ) as table meta [
        Documentation.Name = "fx_LoadPDFDynamic",
        Documentation.LongDescription =
            "Loads all pages from a PDF into a single flattened table. " &
            "Expands up to the specified number of dynamic columns (default = 100), renames them consistently, " &
            "removes null-only columns, and drops page metadata."
    ],

    fx = Value.ReplaceType(fx_LoadPDFDynamic, fx_type)
in
    fx
