SELECT DISTINCT [parksIDs]![CarNo] & " - " & [parksIDs]![offLastName] AS Expr1, parksIDs.Isv, parksIDs.Avail, parksIDs.pId
FROM parksIDs
WHERE (((parksIDs.Isv)=1))
ORDER BY parksIDs.Avail DESC