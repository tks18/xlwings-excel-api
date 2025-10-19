---
name: fx_TableUnpivotAuto
category: Table Processing
tags: [table, unpivot, reshape, normalize, autodetect]
description: "Unpivot selected data type columns automatically (defaults to numeric types)."
version: "v2.0"
---
let
    fx_TableUnpivotAuto =
        (
            sourceTable as table,
            optional dataTypesToUnpivot as list,   // e.g. {type number, type date}
            optional keyColumns as list             // e.g. {"CustomerID"}
        ) 
        as table =>
        let
            // Step 1: Define default key columns (first column if not provided)
            defaultKeyCols = if (keyColumns <> null) then keyColumns else {List.First(Table.ColumnNames(sourceTable))},

            // Step 2: Define default data types to unpivot (numeric types)
            defaultTypes = {
                type number, 
                Int64.Type, 
                Double.Type, 
                Currency.Type, 
                Percentage.Type
            },
            typesToUse = if (dataTypesToUnpivot <> null) then dataTypesToUnpivot else defaultTypes,

            // Step 3: Detect column types
            colTypes = Table.Schema(sourceTable)[Kind],
            colNames = Table.ColumnNames(sourceTable),
            colTypePairs = List.Zip({colNames, colTypes}),

            // Step 4: Identify columns matching allowed data types
            numericCols = 
                List.Transform(
                    List.Select(
                        colTypePairs, 
                        each List.Contains(
                            typesToUse, 
                            Value.Type(Table.Column(sourceTable, _{0}){0}) 
                        )
                    ), each _{0}
                ),

            // Step 5: Fallback if detection fails — use Table.Schema Kind field
            fallbackNumericCols =
                List.Select(
                    Table.Schema(sourceTable)[Name],
                    (colName) =>
                        let
                            kind = Table.Schema(sourceTable){[Name=colName]}[Kind]
                        in
                            List.Contains({"number", "int64", "double", "currency", "percentage"}, Text.Lower(kind))
                ),

            // Step 6: Final columns to unpivot
            colsToUnpivot = if List.Count(numericCols) > 0 then numericCols else fallbackNumericCols,
            unpivotOtherColList = List.Union({ defaultKeyCols, List.Difference(colNames, colsToUnpivot) }),

            // Step 7: Unpivot
            result = Table.UnpivotOtherColumns(sourceTable, unpivotOtherColList, "Attribute", "Value")
        in
            result,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "Input table to unpivot automatically."
        ]),
        optional dataTypesToUnpivot as (type list meta [
            Documentation.FieldCaption = "Data Types to Unpivot",
            Documentation.FieldDescription = "List of data types to unpivot (defaults to numeric)."
        ]),
        optional keyColumns as (type list meta [
            Documentation.FieldCaption = "Key Columns",
            Documentation.FieldDescription = "List of key columns to keep (defaults to first column)."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableUnpivotAuto",
        Documentation.LongDescription =
            "Automatically unpivots all numeric-type columns by default. You can specify custom data types or key columns as needed."
    ],

    fx = Value.ReplaceType(fx_TableUnpivotAuto, fx_type)
in
    fx
