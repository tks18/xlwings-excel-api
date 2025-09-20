---
name: fx_TableRemoveDuplicateRows
category: Table
tags: [table, clean, deduplicate, performance]
description: "Remove duplicate rows based on all columns."
version: "v1.0"
---
let
    fx_TableRemoveDuplicateRows = 
        (sourceTable as table) as table =>
        Table.Distinct(sourceTable),

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "Input table from which duplicate rows should be removed."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableRemoveDuplicateRows",
        Documentation.LongDescription =
            "Removes duplicate rows based on all columns. Uses Table.Distinct for optimal performance."
    ],

    fx = Value.ReplaceType(fx_TableRemoveDuplicateRows, fx_type)
in
    fx