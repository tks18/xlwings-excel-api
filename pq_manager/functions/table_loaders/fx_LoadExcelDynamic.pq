---
name: fx_LoadExcelDynamic
category: Table Loaders
tags: [excel, sheet, table, import, dynamic, flexible]
description: "Load a sheet or table from an Excel file, binary, or already-loaded workbook."
version: "v1.0"
---
let
    fx_LoadExcelDynamic =
        (
            source as any,              // File path, Binary, or Excel workbook
            sourceType as text,         // "Sheet" or "Table"
            itemName as text            // Name of the Sheet or Table
        )
        as table =>
        let
            // Step 1: Handle source type
            sourceWorkbook =
                if Value.Is(source, type binary) then
                    Excel.Workbook(source, true)
                else if Value.Is(source, type text) then
                    Excel.Workbook(File.Contents(source), true)
                else if Table.HasColumns(source, {"Name", "Data"}) then
                    source
                else
                    error "Invalid input. Provide a valid file path, binary, or Excel.Workbook output.",

            // Step 2: Normalize input type (case-insensitive)
            sourceTypeNormalized = Text.Lower(sourceType),

            // Step 3: Filter for the requested item
            filteredItem =
                if sourceTypeNormalized = "sheet" then
                    Table.SelectRows(sourceWorkbook, each [Kind] = "Sheet" and [Name] = itemName)
                else if sourceTypeNormalized = "table" then
                    Table.SelectRows(sourceWorkbook, each [Kind] = "Table" and [Name] = itemName)
                else
                    error "Invalid sourceType. Use 'Sheet' or 'Table'.",

            // Step 4: Extract data
            output =
                if Table.RowCount(filteredItem) = 0 then
                    error Text.Format("No item found with name of type")
                else
                    filteredItem{0}[Data]
        in
            output,

    fx_type = type function (
        source as (type any meta [
            Documentation.FieldCaption = "Excel Source",
            Documentation.FieldDescription = "Can be a file path (text), a binary (File.Contents output), or an already loaded Excel.Workbook table."
        ]),
        sourceType as (type text meta [
            Documentation.FieldCaption = "Source Type",
            Documentation.FieldDescription = "Either 'Sheet' or 'Table'."
        ]),
        itemName as (type text meta [
            Documentation.FieldCaption = "Item Name",
            Documentation.FieldDescription = "The name of the sheet or table to load."
        ])
    ) as table meta [
        Documentation.Name = "fx_LoadExcelDynamic",
        Documentation.LongDescription =
            "Loads data dynamically from an Excel sheet or table, whether the source is a file path, binary, or an already-loaded Excel.Workbook table."
    ],

    fx = Value.ReplaceType(fx_LoadExcelDynamic, fx_type)
in
    fx
