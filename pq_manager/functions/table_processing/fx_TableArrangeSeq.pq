---
name: fx_TableArrangeByPattern
category: Table Processing
tags: [table, arrange, columns, order, pattern]
description: "Arrange columns in numerical order based on a specified prefix pattern (e.g., 'Col_1', 'Col_2', ...)."
version: "v1.1"
---
let
    fx_TableArrangeByPattern =
        (
            sourceTable as table,
            pattern as text,
            optional order as text
        ) as table =>
        let
            // normalize order
            sortOrder = if order <> null and Text.Upper(order) = "DESC" then "DESC" else "ASC",

            // all column names from source
            colNames = Table.ColumnNames(sourceTable),

            // columns that start with the pattern
            matchingCols = List.Select(colNames, each Text.StartsWith(_, pattern)),

            // helper: try extract number after pattern and convert to number, else null
            extractNumber = (colName as text) as nullable number =>
                let
                    suffix = try Text.AfterDelimiter(colName, pattern) otherwise null,
                    num = if suffix <> null then try Number.From(Text.Trim(suffix)) otherwise null else null
                in
                    num,

            // build list of {colName, number} for those that yielded a valid number
            pairs =
                List.Transform(
                    matchingCols,
                    each {_, extractNumber(_)}
                ),

            // keep only pairs where number is not null (safe)
            numericPairs = List.Select(pairs, each _{1} <> null),

            // sort numericPairs by the numeric value (ascending or descending)
            sortedPairs =
                if List.Count(numericPairs) = 0 then
                    {}
                else if sortOrder = "DESC" then
                    List.Sort(numericPairs, (a, b) => if a{1} < b{1} then 1 else if a{1} > b{1} then -1 else 0)
                else
                    List.Sort(numericPairs, (a, b) => if a{1} < b{1} then -1 else if a{1} > b{1} then 1 else 0),

            // extract sorted column names (in numeric order)
            sortedMatchingColNames = List.Transform(sortedPairs, each _{0}),

            // preserve original-order non-matching columns (columns that weren't in sortedMatchingColNames)
            nonMatchingCols = List.Select(colNames, each not List.Contains(sortedMatchingColNames, _)),

            // final column order: sorted matching columns first, then the non-matching columns (preserves their original order)
            finalColumnOrder = List.Combine({sortedMatchingColNames, nonMatchingCols}),

            // if no numeric matches found, return source table as-is
            result =
                if List.Count(sortedMatchingColNames) = 0 then
                    sourceTable
                else
                    Table.ReorderColumns(sourceTable, finalColumnOrder, MissingField.Ignore)
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
            "Example: pattern 'Col_' will reorder ['Col_3','Col_1','Col_2','Name'] to ['Col_1','Col_2','Col_3','Name'] when order='ASC'.\n" &
            "Non-matching columns (or columns with non-numeric suffixes) are appended after the sorted matching columns in their original order."
    ],

    fx = Value.ReplaceType(fx_TableArrangeByPattern, fx_type)
in
    fx
