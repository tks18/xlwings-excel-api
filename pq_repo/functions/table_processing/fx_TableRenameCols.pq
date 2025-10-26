---
name: fx_TableRenameSequential
category: Table Processing
tags: [table, rename, columns, sequential]
description: "Rename all columns sequentially as col_1, col_2, ... col_n."
version: "v1.0"
---
let
    fx_TableRenameSequential =
        (sourceTable as table, optional prefix as text) as table =>
        let
            // Use "col_" as default prefix if not provided
            colPrefix = if prefix <> null then prefix else "col_",

            colNames = Table.ColumnNames(sourceTable),
            colCount = List.Count(colNames),

            newNames = List.Transform({1..colCount}, each colPrefix & Text.From(_)),
            renamePairs = List.Zip({colNames, newNames}),
            result = Table.RenameColumns(sourceTable, renamePairs, MissingField.Ignore)
        in
            result,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "The input table whose columns will be renamed sequentially."
        ]),
        optional prefix as (type text meta [
            Documentation.FieldCaption = "Prefix",
            Documentation.FieldDescription = "Optional text prefix to use for new column names (default = 'Col ')."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableRenameSequential",
        Documentation.LongDescription =
            "This function renames all columns in the input table sequentially.\n" &
            "By default, columns are renamed as 'Col 1', 'Col 2', ... 'Col n'.\n" &
            "Optionally, you can provide a custom prefix (e.g., 'Field ' → 'Field 1', 'Field 2', …)."
    ],

    fx = Value.ReplaceType(fx_TableRenameSequential, fx_type)
in
    fx