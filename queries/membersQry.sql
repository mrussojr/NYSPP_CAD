SELECT parksIDs.offLastName, parksIDs.offFirstName, parksIDs.pId, parksIDs.pZone, parksIDs.Active
FROM parksIDs
WHERE (((parksIDs.pZone)=[Forms]![SearchEvents]![zone]) AND ((parksIDs.Active)=1))
ORDER BY parksIDs.offLastName