---
name: fx_TableUnpivotAuto
category: Table
tags: [table, unpivot, reshape, normalize]
description: "Unpivot all non-key columns automatically."
version: "v1.0"
---
let
    fx_TableUnpivotAuto =
        (sourceTable as table) as table =>
        let
            keyCols = List.FirstN(Table.ColumnNames(sourceTable),1),
            result = Table.UnpivotOtherColumns(sourceTable, keyCols, "Attribute", "Value")
        in
            result,

    fx_type = type function (
        sourceTable as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "Input table to unpivot automatically."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableUnpivotAuto",
        Documentation.LongDescription =
            "Unpivots all columns except the first column into Attribute-Value pairs."
    ],

    fx = Value.ReplaceType(fx_TableUnpivotAuto, fx_type)
in
    fx
