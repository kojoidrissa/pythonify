I'm working on a project to find exceptions between two lists of items. Specifically, for situations where there's a Feature Code or Cost Center in GEAC that's not used in the budget.

This **SHOULD** help with the monthly budget updates, but it could be useful in a lot of other situations as well.

##This project has changed goals. Now the goals are to:

-  Examine methods to 'update' spreadsheets, probably by reading them, making adjustments to certain cells and re-writing the entire workbook. This seems to be the only viable way to "update" a spreadsheet. If it works as I expect, it SHOULD have the advantage of maintaining links b/w cells & sheets by copying the **entire** cell object, not just the value.
-  Once the above issue has been worked out, the plan is to generate two things:
    -  a list of Feature Codes that MATCH what's in GEAC (acutual or budget), so the spreadsheet can be updated
    -  a list of Feature Code 'exceptions' (along with their values) that are in GEAC, but NOT in the spreadsheet being updated. These are the FCs that will have to be updated manually. These should appear in a seperate sheet in one of the following formats (depending on what's easier/faster/better):
        -  One Cost Center/Column, with the exception Feature Codes beneath it and their values to the right of the FC
        -  one exception/row, with each row containing:
            -  Cost Center; FC; Value
