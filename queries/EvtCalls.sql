SELECT EventCallAssoc.EventID, Year([Date1]) AS Expr3, commlog.Date1, commlog.Time1, commlog.UnitCalled, commlog.SourceCall, commlog.Reason, commlog.Narrative, commlog.Dispatcher, Left([Date1],Len([Date1])-5) AS Expr1, TimeValue([Time1]) AS Expr2
FROM commlog INNER JOIN EventCallAssoc ON commlog.id = EventCallAssoc.CallID
WHERE (((EventCallAssoc.EventID)=[Forms]![ViewEvent]![Text25]))
ORDER BY Year([Date1]) DESC , commlog.Date1 DESC , commlog.Time1 DESC