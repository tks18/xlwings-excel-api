---
name: fx_TableArrangeByKeywordOrder
category: Table Processing
tags:
- table
- reorder
- columns
- keyword
- pattern
dependencies: []
description: Reorders specific columns within each group extracted using a supplied pattern. Keyword-to-order map used to sort inside groups. Pattern may be any text containing 'group' and 'subgroup'.
version: V1.0
---

let
    fx_TableArrangeByKeywordOrder =
        (tbl as table, orderList as list, colsToArrange as list, pattern as text) as table =>
        let
            // --- Step 1: Extract placeholders from pattern (e.g. ["group", "subgroup"])
            placeholderRegex = Text.RegexReplace(pattern, "[\\[\\]]", ""),  // remove [ and ]
            parts = Text.Split(placeholderRegex, "."),
            
            // --- Step 2: Extract all keywords from orderList for quick lookup
            keywordList = List.Transform(orderList, each Text.Lower(_{0})),
            
            // --- Step 3: Build lookup for order mapping
            orderMap = Record.FromList(List.Transform(orderList, each _{1}), List.Transform(orderList, each Text.Lower(_{0}))),
            
            // --- Step 4: Extract group & subgroup names from column names
            // Convert [A].[B] to record [group=A, subgroup=B]
            colMap =
                List.Transform(
                    colsToArrange,
                    (colName) =>
                        let
                            cleanName = Text.Remove(colName, {"[", "]"}),
                            splitName = Text.Split(cleanName, "."),
                            rec = [ group = if List.Count(splitName) >= 1 then splitName{0} else null,
                                    subgroup = if List.Count(splitName) >= 2 then splitName{1} else null ],
                            subgroupKey = Text.Lower(rec[subgroup]),
                            orderVal = try Record.Field(orderMap, subgroupKey) otherwise null
                        in
                            [ Column = colName, Group = rec[group], Subgroup = rec[subgroup], Order = orderVal ]
                ),

            // --- Step 5: Filter only those columns that have a matching order
            filteredCols = List.Select(colMap, each _[Order] <> null),

            // --- Step 6: Sort columns by (Group, Order)
            sortedCols = List.Sort(filteredCols, (x, y) =>
                if x[Group] = y[Group] then Value.Compare(x[Order], y[Order])
                else Value.Compare(x[Group], y[Group])
            ),

            // --- Step 7: Extract final column order
            finalOrder = List.Combine({
                List.RemoveItems(Table.ColumnNames(tbl), colsToArrange),
                List.Transform(sortedCols, each _[Column])
            }),

            // --- Step 8: Reorder table columns
            finalTable = Table.ReorderColumns(tbl, finalOrder, MissingField.Ignore)
        in
            finalTable,

    fx_type = type function (
        tbl as (type table meta [
            Documentation.FieldCaption = "Source Table",
            Documentation.FieldDescription = "The input table where columns will be reordered inside groups."
        ]),
        orderList as (type list meta [
            Documentation.FieldCaption = "Keyword Order Map",
            Documentation.FieldDescription = "List of {Keyword, Order} pairs e.g. {{'Client Submitted',1},{'Total',2}}"
        ]),
        colsToArrange as (type list meta [
            Documentation.FieldCaption = "Columns to Rearrange",
            Documentation.FieldDescription = "Exact column names that should be considered for rearrangement inside their group."
        ]),
        pattern as (type text meta [
            Documentation.FieldCaption = "Pattern",
            Documentation.FieldDescription = "Arbitrary pattern containing the tokens 'group' and 'subgroup' to indicate how to parse column names (e.g., '[group].[subgroup]')."
        ])
    ) as table meta [
        Documentation.Name = "fx_TableArrangeByKeywordOrder",
        Documentation.LongDescription =
            "Reorders only the specified columns inside each group extracted dynamically using the 'pattern'. " &
            "The pattern may contain any text but must include the tokens 'group' and 'subgroup' to indicate extraction points. " &
            "Columns not in colsToRearrange are left untouched. Keyword matching is case-insensitive substring match."
    ],

    fx = Value.ReplaceType(fx_TableArrangeByKeywordOrder, fx_type)
in
    fx