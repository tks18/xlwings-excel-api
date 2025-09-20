---
name: fx_TableKeepColumns
category: Table
tags: [table, select, performance, dynamic]
description: "Keep only the specified columns from a table."
version: "v1.0"
---
let
    fx_TableKeepColumns = 
        (sourceTable as table, columnsToKeep as list) as table =>
        let
            existing = List.Intersect({Table.ColumnNames(sourceTable), columnsToKeep}),
            result = Table.SelectColumns(sourceTable, existing)
        in
            result,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "Input table to filter columns from."
        ]),
        columnsToKeep as (type list meta [
            Documentation.FieldCaption = "Columns to Keep",
            Documentation.FieldDescription = "List of column names that should be retained."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableKeepColumns",
        Documentation.LongDescription =
            "Selectively keeps only the columns listed in columnsToKeep, ignoring missing columns."
    ],

    fx = Value.ReplaceType(fx_TableKeepColumns, fx_type)
in
    fx
