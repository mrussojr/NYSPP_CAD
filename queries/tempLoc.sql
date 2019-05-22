SELECT locations.Location, locations.ParksID, Right([Location],Len([Location])-2) AS Expr1, locations.Zone
FROM locations
WHERE (((locations.ParksID) Is Null))
ORDER BY Right([Location],Len([Location])-2)