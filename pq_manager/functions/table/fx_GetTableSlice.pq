---
name: fx_TableSliceDynamic
category: Table
tags: [table, slice, dynamic, performance]
description: "Slice a table from start to end row based on flexible search criteria."
version: "v1.0"
---
let
    fx_TableSliceDynamic =
        (
            sourceTable as table,
            startCols as list,
            startValues as list,
            startOffset as number,
            endCols as list,
            endValues as list,
            endOffset as number,
            matchMode as text,       // "Exact" or "Contains"
            priorityMode as text     // "FirstFound" or "LastFound"
        ) as table =>

        let
            rowCount = Table.RowCount(sourceTable),

            // Helper to find row index
            findRow = (cols as list, values as list, offset as number, defaultRow as number, searchMode as text, priority as text) as number =>
                let
                    // Create list of all matches for each column/value
                    allMatches =
                        List.Combine(
                            List.Transform(cols, (col) =>
                                List.Transform(values, (v) =>
                                    let
                                        colValues = Table.Column(sourceTable, col),
                                        idx =
                                            if searchMode = "Contains" then
                                                List.PositionOfAny(List.Transform(colValues, Text.From), {Text.From(v)}, Occurrence.First)
                                            else
                                                List.PositionOf(colValues, v)
                                    in
                                        if idx = -1 then null else idx + offset
                                )
                            )
                        ),
                    validMatches = List.RemoveNulls(allMatches),
                    selectedRow =
                        if List.IsEmpty(validMatches) then defaultRow
                        else if priority = "FirstFound" then List.Min(validMatches)
                        else List.Max(validMatches)
                in
                    selectedRow,

            startRow = findRow(startCols, startValues, startOffset, 0, matchMode, priorityMode),
            endRow   = findRow(endCols, endValues, endOffset, rowCount-1, matchMode, priorityMode),

            // Ensure bounds
            startRowAdj = Number.Max(0, startRow),
            endRowAdj = Number.Min(rowCount - 1, endRow),

            result =
                if startRowAdj > endRowAdj then
                    #table(Table.ColumnNames(sourceTable), {}) // empty table if invalid range
                else
                    Table.FirstN(
                        Table.Skip(sourceTable, startRowAdj),
                        endRowAdj - startRowAdj + 1
                    )
        in
            result,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption="Source Table",
            Documentation.FieldDescription="Input table to slice."
        ]),
        startCols as (type list meta [
            Documentation.FieldCaption="Start Columns",
            Documentation.FieldDescription="Columns to search for start marker."
        ]),
        startValues as (type list meta [
            Documentation.FieldCaption="Start Values",
            Documentation.FieldDescription="Values to match (priority order) for start."
        ]),
        startOffset as (type number meta [
            Documentation.FieldCaption="Start Offset",
            Documentation.FieldDescription="Adjust starting row (can be negative)."
        ]),
        endCols as (type list meta [
            Documentation.FieldCaption="End Columns",
            Documentation.FieldDescription="Columns to search for end marker."
        ]),
        endValues as (type list meta [
            Documentation.FieldCaption="End Values",
            Documentation.FieldDescription="Values to match (priority order) for end."
        ]),
        endOffset as (type number meta [
            Documentation.FieldCaption="End Offset",
            Documentation.FieldDescription="Adjust ending row (can be negative)."
        ]),
        matchMode as (type text meta [
            Documentation.FieldCaption="Match Mode",
            Documentation.FieldDescription="Exact or Contains."
        ]),
        priorityMode as (type text meta [
            Documentation.FieldCaption="Priority Mode",
            Documentation.FieldDescription="FirstFound or LastFound."
        ])
    ) as table meta [
        Documentation.Name="fx_TableSliceDynamic",
        Documentation.LongDescription="Slices a table from start row to end row based on dynamic search criteria, offsets, and priority settings."
    ],

    fx = Value.ReplaceType(fx_TableSliceDynamic, fx_type)
in
    fx
