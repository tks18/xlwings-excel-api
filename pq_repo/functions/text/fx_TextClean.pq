---
name: fx_TextClean
category: Text Processing
tags: [text, clean, trim, whitespace, normalize]
description: "Fast text cleaning function that removes non-printable characters, trims, removes extra spaces, and normalizes whitespace."
version: "v1.0"
---
let
    fx_TextClean =
        (
            inputText as nullable text
        ) as nullable text =>
        let
            // Return null safely
            result =
                if inputText = null then
                    null
                else
                    let
                        // Ensure text type
                        txt = Text.From(inputText),

                        // Remove non-printable characters (fast built-in)
                        cleaned = Text.Clean(txt),

                        // Replace common whitespace characters with space
                        normalizedWhitespace =
                            Text.Replace(
                                Text.Replace(
                                    Text.Replace(cleaned, "#(lf)", " "),
                                "#(cr)", " "),
                            "#(tab)", " "
                            ),

                        // Trim leading/trailing spaces
                        trimmed = Text.Trim(normalizedWhitespace),

                        // Collapse multiple spaces in a single pass
                        splitList = Text.Split(trimmed, " "),
                        nonEmpty = List.Select(splitList, each _ <> ""),
                        singleSpaced = Text.Combine(nonEmpty, " ")
                    in
                        singleSpaced
        in
            result,

    fx_type = type function (
        inputText as (type nullable text meta [
            Documentation.FieldCaption = "Input Text",
            Documentation.FieldDescription = "The text string to clean and normalize."
        ])
    ) as nullable text meta [
        Documentation.Name = "fx_TextClean",
        Documentation.LongDescription =
            "This function efficiently cleans input text by:\n" &
            "- Removing non-printable characters\n" &
            "- Replacing line breaks and tabs with spaces\n" &
            "- Trimming leading and trailing spaces\n" &
            "- Collapsing multiple consecutive spaces into a single space\n" &
            "Designed for high-performance bulk text transformations."
    ],

    fx = Value.ReplaceType(fx_TextClean, fx_type)
in
    fx
