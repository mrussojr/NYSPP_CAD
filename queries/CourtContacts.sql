SELECT Courts.*, Justices.*
FROM Courts LEFT JOIN Justices ON Courts.ID = Justices.CourtID