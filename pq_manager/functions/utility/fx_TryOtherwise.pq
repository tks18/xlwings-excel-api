---
name: fx_TryOtherwise
category: Utility
tags: [error-handling, safety, try, otherwise, fallback]
description: "Safely executes a function with given input. Returns fallback value if an error occurs."
version: "v1.0"
---
let
    fx_TryOtherwise =
        (operation as function, optional alternate as any) as any =>
        let
            // Evaluate operation with input safely
            attempt = try operation(),

            // Check for error
            result = if attempt[HasError] 
                     then if alternate <> null then alternate else null 
                     else attempt[Value]
        in
            result,

    fx_type = type function (
        operation as (type function meta [
            Documentation.FieldCaption = "Function to Execute",
            Documentation.FieldDescription = "The function that will be executed safely. Example: Number.FromText."
        ]),
        optional alternate as (type any meta [
            Documentation.FieldCaption = "Fallback Value",
            Documentation.FieldDescription = "Optional value to return if the function execution results in an error (default = null)."
        ])
    ) as any meta [
        Documentation.Name = "fx_TryOtherwise",
        Documentation.LongDescription =
            "This function wraps another function with built-in error handling.\n" &
            "If the operation succeeds, it returns the result.\n" &
            "If an error occurs, it returns the provided fallback value (or null if not provided).\n\n" &
            "Examples:\n" &
            "• fx_TryOtherwise(Number.FromText, \"123\", -1) → 123\n" &
            "• fx_TryOtherwise(Number.FromText, \"abc\", -1) → -1\n" &
            "• fx_TryOtherwise(Text.Length, null, 0) → 0"
    ],

    fx = Value.ReplaceType(fx_TryOtherwise, fx_type)
in
    fx
