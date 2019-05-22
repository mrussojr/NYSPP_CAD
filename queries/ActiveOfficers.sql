SELECT parksIDs.offLastName, parksIDs.offFirstName, parksIDs.offShield, parksIDs.pId, parksIDs.pZone, parksIDs.Isv, parksIDs.Avail, parksIDs.CarNo, parksIDs.Active
FROM parksIDs
WHERE (((parksIDs.Active)=1))
ORDER BY parksIDs.offLastName