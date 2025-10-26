---
name: fx_TableSlice
category: Table Processing
tags: [table, slice, dynamic, performance]
description: "Slice a table from start to end row based on flexible search criteria (startCols/startValues/endCols/endValues), supports Exact/Contains and FirstFound/LastFound."
version: "v2.0"
---
let
    fx_TableSlice =
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
            // basic inputs
            rowCount = Table.RowCount(sourceTable),
            colNames = Table.ColumnNames(sourceTable),

            // ensure provided column lists are valid (filter out missing columns)
            validStartCols = List.Intersect({startCols, colNames}),
            validEndCols = List.Intersect({endCols, colNames}),

            // add index once and buffer for repeated searches
            indexed = Table.AddIndexColumn(sourceTable, "__idx", 0, 1),
            indexedBuf = Table.Buffer(indexed),

            // function: returns list of matching __idx for a given value across provided columns
            getIndicesForValue = (value as any, cols as list) as list =>
                let
                    colsToUse = List.Intersect({cols, Table.ColumnNames(indexedBuf)}),
                    // for each column produce the rows that match the value
                    matchingTables =
                        List.Transform(colsToUse, (c) =>
                            Table.SelectRows(indexedBuf, each
                                let
                                    cell = try Record.Field(_, c) otherwise null,
                                    // Contains: case-insensitive on textified values
                                    containsCond = try if cell <> null then Text.Contains(Text.From(cell), Text.From(value)) else false otherwise false,
                                    // Exact: try direct equality, fallback to text equality
                                    exactCond = try if cell = value then true else (try Text.From(cell) = Text.From(value) otherwise false) otherwise false
                                in
                                    if Text.Lower(matchMode) = "contains" then containsCond else exactCond
                            )
                        ),
                    combined =
                        if List.IsEmpty(matchingTables)
                        then #table({}, {})    // no valid columns → empty
                        else Table.Combine(matchingTables),
                    idxList = if Table.RowCount(combined) = 0 then {} else List.Distinct(Table.Column(combined, "__idx"))
                in
                    idxList,

            // Resolve start index according to priorityMode
            startIndex =
                if List.IsEmpty(startValues) or List.IsEmpty(validStartCols) then 0
                else if Text.Lower(priorityMode) = "firstfound" then
                    let
                        // for each startValue, get earliest occurrence (min index) across columns; pick the first value that yields a hit (value-priority order)
                        valueEarliestList =
                            List.Transform(startValues, (v) =>
                                let idxs = getIndicesForValue(v, validStartCols) in if List.IsEmpty(idxs) then null else List.Min(idxs)
                            ),
                        nonNull = List.RemoveNulls(valueEarliestList)
                    in
                        if List.IsEmpty(nonNull) then 0 else nonNull{0}
                else
                    // LastFound => gather all indices for all values and pick the latest (max)
                    let
                        allIdxs = List.Combine(List.Transform(startValues, each getIndicesForValue(_, validStartCols)))
                    in
                        if List.IsEmpty(allIdxs) then 0 else List.Max(allIdxs),

            // Resolve end index similarly (default to last row if not found)
            endIndex =
                if List.IsEmpty(endValues) or List.IsEmpty(validEndCols) then rowCount - 1
                else if Text.Lower(priorityMode) = "firstfound" then
                    let
                        valueEarliestList =
                            List.Transform(endValues, (v) =>
                                let idxs = getIndicesForValue(v, validEndCols) in if List.IsEmpty(idxs) then null else List.Min(idxs)
                            ),
                        nonNull = List.RemoveNulls(valueEarliestList)
                    in
                        if List.IsEmpty(nonNull) then rowCount - 1 else nonNull{0}
                else
                    let
                        allIdxs = List.Combine(List.Transform(endValues, each getIndicesForValue(_, validEndCols)))
                    in
                        if List.IsEmpty(allIdxs) then rowCount - 1 else List.Max(allIdxs),

            // apply offsets
            startRowRaw = startIndex + startOffset,
            endRowRaw = endIndex + endOffset,

            // clamp to bounds
            startRow = List.Max({0, List.Min({rowCount - 1, startRowRaw})}),
            endRow = List.Max({0, List.Min({rowCount - 1, endRowRaw})}),

            // result (empty table if invalid range)
            result =
                if startRow > endRow then
                    #table(colNames, {})   // empty table with same columns
                else
                    Table.FirstN(Table.Skip(sourceTable, startRow), endRow - startRow + 1)
        in
            result,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "Input table to slice."
        ]),
        startCols as (type list meta [
            Documentation.FieldCaption = "Start Columns",
            Documentation.FieldDescription = "Columns to search for start marker."
        ]),
        startValues as (type list meta [
            Documentation.FieldCaption = "Start Values",
            Documentation.FieldDescription = "Values to match (priority order) for start."
        ]),
        startOffset as (type number meta [
            Documentation.FieldCaption = "Start Offset",
            Documentation.FieldDescription = "Adjust starting row (can be negative)."
        ]),
        endCols as (type list meta [
            Documentation.FieldCaption = "End Columns",
            Documentation.FieldDescription = "Columns to search for end marker."
        ]),
        endValues as (type list meta [
            Documentation.FieldCaption = "End Values",
            Documentation.FieldDescription = "Values to match (priority order) for end."
        ]),
        endOffset as (type number meta [
            Documentation.FieldCaption = "End Offset",
            Documentation.FieldDescription = "Adjust ending row (can be negative)."
        ]),
        matchMode as (type text meta [
            Documentation.FieldCaption = "Match Mode",
            Documentation.FieldDescription = "Exact or Contains."
        ]),
        priorityMode as (type text meta [
            Documentation.FieldCaption = "Priority Mode",
            Documentation.FieldDescription = "FirstFound or LastFound."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableSlice",
        Documentation.LongDescription =
            "Slices a table from start to end row based on flexible search criteria (startCols/startValues/endCols/endValues). " &
            "Supports Exact/Contains match modes and FirstFound/LastFound priority modes. Uses an indexed buffered table to avoid repeated re-scans."
    ],

    fx = Value.ReplaceType(fx_TableSlice, fx_type)
in
    fx
