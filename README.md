# ExtractPlaceHolders
Extracts placeholders, e.g. {%1} etc., from phrase/sentence to an array.

Example:

Dim arr As variant

    arr = ExtractTerms("This is a test, {%1} of the {%2} terms function")
    
This will extract an array containing {%1} and {%2} in that order.    
