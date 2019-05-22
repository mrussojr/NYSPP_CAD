
"CREATE TABLE [EvtTypes] (
	[ID] Counter
	,[Event_Type] Text (255)
	,[ParkNum] Long )"

"CREATE INDEX [ParkNum] ON [EvtTypes] ([ParkNum]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [EvtTypes] ([ParkNum])  WITH PRIMARY DISALLOW NULL "
