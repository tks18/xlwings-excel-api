---
name: fx_TableRemoveNullCols
category: Table Processing
tags: [table, clean, nulls, performance, metadata]
description: "Remove columns that contain only null values."
version: "v2.0"
---
let
    fx_TableRemoveNullCols = 
        (sourceTable as table) as table =>
        let
            cols = Table.ColumnNames(sourceTable),
            profile = Table.FromRecords(Table.ToRecords(Table.Profile(sourceTable))),
            nonEmpty = Table.SelectRows(profile, each [NullCount] <> [Count])[Column],
            result = Table.SelectColumns(sourceTable, List.Intersect({ cols, nonEmpty }))
        in
            result,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "Input table from which fully null columns will be removed."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableRemoveNullCols",
        Documentation.LongDescription =
            "This function removes all columns that contain only null values." &
            "It is highly efficient because it scans the table only once using Table.Profile."
    ],

    fx = Value.ReplaceType(fx_TableRemoveNullCols, fx_type)
in
    fx