SELECT events.ID, events.Active, events.LastDate, events.LastTime, events.EvtDate, events.OutTime, events.ColorCode, events.IgnoreTimer, IIf(IsNull([LastTime]),DateDiff("n",[Expr3],[Expr5]),DateDiff("n",[Expr4],[Expr5])) AS Expr1, CDate([EvtDate] & " " & [OutTime]) AS Expr3, IIf(IsNull([LastTime]),Null,CDate([LastDate] & " " & [LastTime])) AS Expr4, CDate(Date() & " " & Time()) AS Expr5, events.Location, events.[Off], events.Type, DateDiff("h",[Expr4],[Expr5]) AS Expr6
FROM events
WHERE (((events.Active)=0) AND ((events.IgnoreTimer)=0))
ORDER BY events.LastDate DESC