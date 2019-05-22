
"CREATE TABLE [parksIDs] (
	[id] Counter
	,[offLastName] Text (255)
	,[offFirstName] Text (255)
	,[offShield] Integer
	,[pId] Integer
	,[pDistrict] Long
	,[pZone] Integer
	,[pCounty] Long
	,[pStation] Long
	,[Isv] Long
	,[Avail] Long
	,[AssId] Long
	,[SecondaryId] Long
	,[CarNo] Text (255)
	,[Active] Long
	,[StartTime] Text (255) )"

"CREATE INDEX [AssId] ON [parksIDs] ([StartTime]) "

"CREATE INDEX [id] ON [parksIDs] ([StartTime]) "

"CREATE INDEX [pId] ON [parksIDs] ([StartTime]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [parksIDs] ([StartTime])  WITH PRIMARY DISALLOW NULL "

"CREATE INDEX [SecondaryId] ON [parksIDs] ([StartTime]) "
