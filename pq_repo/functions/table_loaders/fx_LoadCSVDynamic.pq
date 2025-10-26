---
name: fx_LoadCsvDynamic
category: Table Loaders
tags: [csv, dynamic, import, columns, table]
description: "Load a CSV file with dynamic column handling and proper quote support."
version: "v2.0"
---
let
    fx_LoadCsvDynamic =
        (filePath as text, optional delimiter as text, optional encoding as number, optional hasHeaders as logical) as table =>
        let
            sep = if delimiter <> null then delimiter else ",",
            enc = if encoding <> null then encoding else 65001,
            headers = if hasHeaders <> null then hasHeaders else true,

            // Use Csv.Document which handles quotes properly
            csv = Csv.Document(
                File.Contents(filePath),
                [Delimiter = sep, Encoding = enc, QuoteStyle = QuoteStyle.Csv]
            ),

            // Promote headers if needed
            final = if headers then Table.PromoteHeaders(csv, [PromoteAllScalars=true]) else csv
        in
            final,

    fx_type = type function (
        filePath as (type text meta [
            Documentation.FieldCaption = "CSV File Path",
            Documentation.FieldDescription = "Path to the CSV file."
        ]),
        optional delimiter as (type text meta [
            Documentation.FieldCaption = "Delimiter",
            Documentation.FieldDescription = "Optional delimiter (default = ',')."
        ]),
        optional encoding as (type number meta [
            Documentation.FieldCaption = "Encoding",
            Documentation.FieldDescription = "Optional file encoding (default = UTF-8)."
        ]),
        optional hasHeaders as (type logical meta [
            Documentation.FieldCaption = "Has Headers",
            Documentation.FieldDescription = "Indicates if CSV has headers (default = true)."
        ])
    ) as table meta [
        Documentation.Name = "fx_LoadCsvDynamic",
        Documentation.LongDescription =
            "Loads a CSV file with dynamic columns and proper quote handling. " &
            "Automatically promotes headers if present. " &
            "Supports custom delimiter and encoding."
    ],

    fx = Value.ReplaceType(fx_LoadCsvDynamic, fx_type)
in
    fx
