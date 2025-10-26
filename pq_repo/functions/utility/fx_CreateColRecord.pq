---
name: fx_CreateColRecord
category: Utility
tags: [record, column, dynamic, addcolumn, metadata]
description: "Creates a record defining a column with Name, Type, and Function fields for dynamic table transformations."
version: "v1.0"
---
let
    fx_CreateColRecord =
        (
            columnName as text,
            columnType as type,
            columnFunction as function
        ) as record =>
        let
            recordOut = [
                Name = columnName,
                Type = columnType,
                Function = columnFunction
            ]
        in
            recordOut,

    fx_type = type function (
        columnName as (type text meta [
            Documentation.FieldCaption = "Column Name",
            Documentation.FieldDescription = "The name of the column to define."
        ]),
        columnType as (type type meta [
            Documentation.FieldCaption = "Column Type",
            Documentation.FieldDescription = "The Power Query type of the column (e.g., type text, type number)."
        ]),
        columnFunction as (type function meta [
            Documentation.FieldCaption = "Column Function",
            Documentation.FieldDescription = "A row-level function defining how to compute the column value."
        ])
    ) as record meta [
        Documentation.Name = "fx_CreateColRecord",
        Documentation.LongDescription =
            "This function returns a record with three fields: Name, Type, and Function.\n" &
            "It is useful for building lists of column definitions that can be passed to other functions such as fx_AddColumnsFromRecords.\n"
    ],

    fx = Value.ReplaceType(fx_CreateColRecord, fx_type)
in
    fx
