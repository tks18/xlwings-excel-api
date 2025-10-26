---
name: fx_TableAddColsDynamic
category: Table Processing
tags: [table, addcolumn, dynamic, record, automation]
description: "Add multiple columns to a table dynamically based on a list of records defining name, type, and function."
version: "v1.0"
---
let
    fx_TableAddColsDynamic =
        (
            sourceTable as table,
            columnsToAdd as list     // each item: [Name="...", Type=type text, Function=(r)=>...]
        )
        as table =>
        let
            addColumns =
                List.Accumulate(
                    columnsToAdd,
                    sourceTable,
                    (state, colRec) =>
                        let
                            colName = Record.FieldOrDefault(colRec, "Name", null),
                            colType = Record.FieldOrDefault(colRec, "Type", null),
                            colFunc = Record.FieldOrDefault(colRec, "Function", null),
                            newTable =
                                if colName <> null and Value.Is(colFunc, type function) then
                                    if colType <> null then
                                        Table.AddColumn(state, colName, colFunc, colType)
                                    else
                                        Table.AddColumn(state, colName, colFunc)
                                else
                                    state
                        in
                            newTable
                )
        in
            addColumns,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "The input table to which columns will be added."
        ]),
        columnsToAdd as (type list meta [
            Documentation.FieldCaption = "Columns To Add",
            Documentation.FieldDescription = "A list of records defining each column with Name, Type, and Function keys."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableAddColsDynamic",
        Documentation.LongDescription =
            "Adds multiple columns dynamically. Each record in the list should define: [Name = text, Type = type, Function = function(row) => ...]"
    ],

    fx = Value.ReplaceType(fx_TableAddColsDynamic, fx_type)
in
    fx
