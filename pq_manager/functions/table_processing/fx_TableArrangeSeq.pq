---
name: fx_TableArrangeByPattern
category: Table Processing
tags: [table, arrange, columns, order, pattern]
description: "Arrange columns in numerical order based on a specified prefix pattern (e.g., 'Col_1', 'Col_2', ...)."
version: "v1.0"
---
let
    fx_TableArrangeByPattern =
        (
            sourceTable as table,
            pattern as text,
            optional order as text
        ) as table =>
        let
            // Default order is ascending
            sortOrder = if order <> null and Text.Upper(order) = "DESC" then "DESC" else "ASC",

            // Get all column names
            colNames = Table.ColumnNames(sourceTable),

            // Filter columns matching the given pattern
            matchingCols = List.Select(colNames, each Text.StartsWith(_, pattern)),

            // Extract numeric part after the pattern (e.g., "Col_5" → 5)
            colNumbers = List.Transform(matchingCols, each 
                try Number.From(Text.AfterDelimiter(_, pattern)) 
                otherwise null
            ),

            // Combine names and numbers for sorting
            combined = List.Zip({matchingCols, colNumbers}),

            // Remove any null-numbered columns (safety)
            filtered = List.Select(combined, each _{1} <> null),

            // Sort by number ascending or descending
            sorted = if sortOrder = "DESC" 
                then List.Sort(filtered, Order.Descending, each _{1})
                else List.Sort(filtered, Order.Ascending, each _{1}),

            // Get sorted column names
            sortedCols = List.Transform(sorted, each _{0}),

            // Reorder table columns
            result = Table.ReorderColumns(sourceTable, sortedCols, MissingField.Ignore)
        in
            result,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "The table whose columns will be reordered."
        ]),
        pattern as (type text meta [
            Documentation.FieldCaption = "Column Name Pattern",
            Documentation.FieldDescription = "The common prefix used in column names (e.g., 'Col_')."
        ]),
        optional order as (type text meta [
            Documentation.FieldCaption = "Order",
            Documentation.FieldDescription = "Set 'ASC' for ascending or 'DESC' for descending order (default = 'ASC')."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableArrangeByPattern",
        Documentation.LongDescription =
            "This function arranges table columns numerically based on the numeric suffix extracted after a given pattern.\n" &
            "For example, with pattern 'Col_', columns ['Col_1', 'Col_3', 'Col_2'] will be reordered as ['Col_1', 'Col_2', 'Col_3'].\n" &
            "Optionally specify order = 'DESC' for descending sort."
    ],

    fx = Value.ReplaceType(fx_TableArrangeByPattern, fx_type)
in
    fx
