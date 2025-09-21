---
name: fx_TextSlugify
category: Text
tags: [text, slug, clean, url, normalize]
description: "Convert input text into a URL-friendly slug."
version: "v1.0"
---
let
    fx_TextSlugify =
        (inputText as text, optional delimiter as text) as text =>
        let
            // Default delimiter = "-"
            delim = if delimiter <> null then delimiter else "_",

            // Lowercase
            lower = Text.Lower(inputText),

            // Remove diacritics (accents) – normalize to ASCII
            decomposed = Text.Normalize(lower, TextNormalization.FormD),
            asciiOnly = Text.Select(decomposed, List.Combine({{"a".."z"}, {"0".."9"}, {" "}})),

            // Replace spaces with delimiter
            replaced = Text.Replace(asciiOnly, " ", delim),

            // Collapse multiple delimiters
            collapsed = Text.Combine(List.Select(Text.Split(replaced, delim), each _ <> ""), delim)
        in
            collapsed,

    fx_type = type function (
        inputText as (type text meta [
            Documentation.FieldCaption = "Input Text",
            Documentation.FieldDescription = "The text string to convert into a slug."
        ]),
        optional delimiter as (type text meta [
            Documentation.FieldCaption = "Delimiter",
            Documentation.FieldDescription = "Optional delimiter to use between slug parts (default = '-')."
        ])
    ) as text meta [
        Documentation.Name = "fx_TextSlugify",
        Documentation.LongDescription =
            "Converts the input text into a slug (URL-friendly string)." &
            "- Converts to lowercase" &
            "- Removes accents/diacritics" &
            "- Keeps only letters, numbers, and spaces" &
            "- Replaces spaces with the chosen delimiter (default '_')" &
            "- Collapses multiple delimiters"
    ],

    fx = Value.ReplaceType(fx_TextSlugify, fx_type)
in
    fx
