---
name: fx_LoadJSONDynamic
category: JSON
tags: [json, expand, flatten, file, binary, dynamic]
description: "Load a JSON file (from file path or binary), parse, and recursively expand all nested objects and arrays into a flat table. Allows custom delimiter for nested column names."
version: "v1.1"
---
let
    fx_LoadJSONDynamic =
        (jsonInput as any, optional delimiter as text) as table =>
        let
            // Step 1: Set delimiter (default ".")
            delim = if delimiter = null then "." else delimiter,

            // Step 2: Detect input type (file path, binary, or already-parsed)
            LoadedJSON =
                if Value.Is(jsonInput, type text) then
                    Json.Document(File.Contents(jsonInput))    // file path
                else if Value.Is(jsonInput, type binary) then
                    Json.Document(jsonInput)                   // binary
                else if Value.Is(jsonInput, type table) then
                    jsonInput                                  // already a table
                else
                    jsonInput,                                 // already parsed JSON

            // Step 3: Convert to table if needed
            ToTable = 
                if Type.Is(Value.Type(LoadedJSON), type table) 
                then LoadedJSON
                else if Value.Is(LoadedJSON, type list) then
                    Table.FromList(LoadedJSON, Splitter.SplitByNothing(), {"Value"})
                else
                    Table.FromRecords({LoadedJSON}),

            // Step 4: Recursive expansion
            ExpandRecursively = (tbl as table) as table =>
                let
                    cols = Table.ColumnNames(tbl),
                    expanded = List.Accumulate(
                        cols,
                        tbl,
                        (state, col) =>
                            let
                                firstNonNull = try List.First(List.RemoveNulls(Table.Column(state, col))) otherwise null,
                                newState =
                                    if firstNonNull is record then
                                        Table.ExpandRecordColumn(
                                            state, 
                                            col, 
                                            Record.FieldNames(firstNonNull), 
                                            List.Transform(Record.FieldNames(firstNonNull), each col & delim & _)
                                        )
                                    else if firstNonNull is list then
                                        Table.ExpandListColumn(state, col)
                                    else
                                        state
                            in
                                newState
                    ),
                    stillNested = List.AnyTrue(
                        List.Transform(
                            Table.ColumnNames(expanded), 
                            each List.AnyTrue(
                                List.Transform(Table.Column(expanded, _), 
                                    (v) => (Value.Is(v, type record) or Value.Is(v, type list))
                                )
                            )
                        )
                    ),
                    result = if stillNested then @ExpandRecursively(expanded) else expanded
                in
                    result
        in
            ExpandRecursively(ToTable),

    fx_type = type function (
        jsonInput as (type any meta [
            Documentation.FieldCaption = "JSON Input",
            Documentation.FieldDescription = "File path (text), binary JSON (from Web.Contents, File.Contents, etc.), or already-parsed JSON/table."
        ]),
        optional delimiter as (type text meta [
            Documentation.FieldCaption = "Delimiter",
            Documentation.FieldDescription = "String used to join nested column names (default is '.')."
        ])
    ) as table meta [
        Documentation.Name = "fx_LoadJSONDynamic",
        Documentation.LongDescription = 
            "Loads a JSON file or binary input, parses it into a JSON object, and recursively expands all nested records and lists into a flat table. " &
            "Nested column names are prefixed using Parent{delimiter}Child notation."
    ],

    fx = Value.ReplaceType(fx_LoadJSONDynamic, fx_type)
in
    fx
