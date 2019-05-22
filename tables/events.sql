
"CREATE TABLE [events] (
	[ID] Counter
	,[Shield] Long
	,[Off] Text (255)
	,[OutTime] DateTime
	,[EvtDate] DateTime
	,[Location] Text (255)
	,[Type] Text (255)
	,[Active] Long
	,[LastTime] DateTime
	,[LastDate] DateTime
	,[Narrative] Memo
	,[Address] Text (255)
	,[Cross_Street] Text (255)
	,[MapLocation] Memo
	,[pId] Long
	,[ColorCode] Long
	,[IgnoreTimer] Long
	,[Field1] Text (255) )"

"CREATE INDEX [ColorCode] ON [events] ([Field1]) "

"CREATE INDEX [OId] ON [events] ([Field1]) "

"CREATE INDEX [pId] ON [events] ([Field1]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [events] ([Field1])  WITH PRIMARY DISALLOW NULL "
