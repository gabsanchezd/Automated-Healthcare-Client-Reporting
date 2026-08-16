let
    Source = Folder.Files("C:\Github\Automated Healthcare Client Reporting\data\raw\claims"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"Claim ID", type text}, {"Entity", type text}, {"Patient ID", type text}, {"Date of Service", type date}, {"Provider", type text}, {"Office", type text}, {"Payer", type text}, {"Payer Type", type text}, {"Procedure Code", Int64.Type}, {"Claim Status", type text}, {"Billed Amount", Currency.Type}, {"Allowed Amount", Currency.Type}, {"Initial Payment", Currency.Type}, {"Negotiation Date", type date}, {"IDR Submission Date", type date}, {"Determination Date", type date}, {"Additional Owed", Currency.Type}, {"Additional Collected", Currency.Type}, {"Date Collected", type date}, {"Insurance Balance", Currency.Type}}),
    #"Changed to Text and Date" = Table.TransformColumnTypes(#"Changed Type",{{"Claim ID", type text}, {"Entity", type text}, {"Patient ID", type text}, {"Provider", type text}, {"Office", type text}, {"Payer", type text}, {"Payer Type", type text}, {"Procedure Code", type text}, {"Claim Status", type text}, {"Date of Service", type date}, {"Negotiation Date", type date}, {"IDR Submission Date", type date}, {"Determination Date", type date}, {"Date Collected", type date}, {"Billed Amount", type number}, {"Allowed Amount", type number}, {"Initial Payment", type number}, {"Additional Owed", type number}, {"Additional Collected", type number}, {"Insurance Balance", type number}}),
    #"Trimmed Text" = Table.TransformColumns(#"Changed to Text and Date",{{"Entity", Text.Trim, type text}, {"Provider", Text.Trim, type text}, {"Office", Text.Trim, type text}, {"Payer", Text.Trim, type text}, {"Payer Type", Text.Trim, type text}, {"Claim Status", Text.Trim, type text}}),
    #"Cleaned Text" = Table.TransformColumns(#"Trimmed Text",{{"Entity", Text.Clean, type text}, {"Provider", Text.Clean, type text}, {"Office", Text.Clean, type text}, {"Payer", Text.Clean, type text}, {"Payer Type", Text.Clean, type text}, {"Claim Status", Text.Clean, type text}}),
    #"Merged Queries" = Table.NestedJoin(#"Cleaned Text", {"Claim Status"}, qry_StatusLookup, {"Claim Status"}, "qry_StatusLookup", JoinKind.LeftOuter),
    #"Expanded qry_StatusLookup" = Table.ExpandTableColumn(#"Merged Queries", "qry_StatusLookup", {"Status Group", "Sort Order"}, {"Status Group", "Sort Order"}),
    #"Inserted Year" = Table.AddColumn(#"Expanded qry_StatusLookup", "Year", each Date.Year([Date of Service]), Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Year",{{"Year", "DOS Year"}}),
    #"Inserted Month Name" = Table.AddColumn(#"Renamed Columns", "Month Name", each Date.MonthName([Date of Service]), type text),
    #"Inserted Month" = Table.AddColumn(#"Inserted Month Name", "Month", each Date.Month([Date of Service]), Int64.Type),
    #"Renamed Columns2" = Table.RenameColumns(#"Inserted Month",{{"Month Name", "DOS Month"}, {"Month", "DOS Month Number"}}),
    #"Year-Month" = Table.AddColumn(#"Renamed Columns2", "Year-Month", each Date.ToText([Date of Service], "yyyy-MM")),
    #"Outstanding Amount" = Table.AddColumn(#"Year-Month", "Outstanding Amount", each [Allowed Amount] - [Initial Payment] - [Additional Collected]),
    #"Total Collected" = Table.AddColumn(#"Outstanding Amount", "Total Collected", each [Initial Payment] + [Additional Collected]),
    #"Collection Rate" = Table.AddColumn(#"Total Collected", "Collection Rate", each if [Allowed Amount] = null or [Allowed Amount] = 0
then null
else [Total Collected] / [Allowed Amount]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Collection Rate",{{"Collection Rate", Percentage.Type}, {"Outstanding Amount", type number}, {"Total Collected", type number}}),
    #"Aging Days" = Table.AddColumn(#"Changed Type1", "Aging Days", each if [Date of Service] = null then
    null
else
    try Duration.Days(#date(2025, 12, 31) - Date.From([Date of Service]))
    otherwise null),
    #"Aging Bucket" = Table.AddColumn(#"Aging Days", "Aging Bucket", each if [Aging Days] = null then "Unknown"
else if [Aging Days] < 0 then "Future DOS"
else if [Aging Days] <= 30 then "0-30"
else if [Aging Days] <= 60 then "31-60"
else if [Aging Days] <= 90 then "61-90"
else if [Aging Days] <= 120 then "91-120"
else "120+"),
    #"Missing Claim ID" = Table.AddColumn(#"Aging Bucket", "QC Missing Claim ID", each if [Claim ID] = null or Text.Trim([Claim ID]) = ""
then "ERROR"
else "PASS"),
    #"Missing Entry" = Table.AddColumn(#"Missing Claim ID", "QC Missing Entity", each if [Entity] = null or Text.Trim([Entity]) = ""
then "ERROR"
else "PASS"),
    #"QC Invalid Status" = Table.AddColumn(#"Missing Entry", "QC Invalid Status", each if [Status Group] = null
then "ERROR"
else "PASS"),
    #"QC Negative Billed" = Table.AddColumn(#"QC Invalid Status", "QC Negative Billed", each if [Billed Amount] < 0 then "ERROR" else "PASS"),
    #"QC Negative Collection" = Table.AddColumn(#"QC Negative Billed", "QC Negative Collection", each if [Additional Collected] < 0 then "ERROR" else "PASS"),
    #"QC Allowed vs Billed" = Table.AddColumn(#"QC Negative Collection", "QC Allowed vs Billed", each if [Allowed Amount] > [Billed Amount]
then "REVIEW"
else "PASS"),
    #"QC Collected vs Allowed" = Table.AddColumn(#"QC Allowed vs Billed", "QC Collected vs Allowed", each if [Total Collected] > [Allowed Amount]
then "REVIEW"
else "PASS"),
    #"Balance Difference" = Table.AddColumn(#"QC Collected vs Allowed", "Balance Difference", each [Insurance Balance] - [Outstanding Amount]),
    #"QC Balance Reconciliation" = Table.AddColumn(#"Balance Difference", "QC Balance Reconciliation", each if Number.Abs([Balance Difference]) > 1
then "REVIEW"
else "PASS"),
    #"Overall QC" = Table.AddColumn(#"QC Balance Reconciliation", "Overall QC", each let
    ErrorChecks = List.RemoveNulls({
        [QC Missing Claim ID],
        [QC Missing Entity],
        [QC Invalid Status],
        [QC Negative Billed],
        [QC Negative Collection]
    }),

    ReviewChecks = List.RemoveNulls({
        [QC Allowed vs Billed],
        [QC Collected vs Allowed],
        [QC Balance Reconciliation]
    })
in
    if List.Contains(ErrorChecks, "ERROR") then "ERROR"
    else if List.Contains(ReviewChecks, "REVIEW") then "REVIEW"
    else "PASS")
in
    #"Overall QC"