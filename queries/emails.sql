SELECT AllParkInfo.Park, AllParkInfo.Manager, Left([Manager],InStr(1,[Manager]," ")-1) & "." & Right([Manager],Len([Manager])-InStrRev([Manager]," ")) & "@parks.ny.gov" AS Expr1
FROM AllParkInfo