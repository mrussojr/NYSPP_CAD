SELECT DISTINCT CourtContacts.County, CourtContacts.Jurisdiction
FROM Counties INNER JOIN CourtContacts ON Counties.County = CourtContacts.County