
"CREATE TABLE [Stations] (
	[ID] Counter
	,[Station] Text (255)
	,[StationPID] Long )"

"CREATE INDEX [ID] ON [Stations] ([StationPID]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [Stations] ([StationPID])  WITH PRIMARY DISALLOW NULL "

"CREATE INDEX [StationPID] ON [Stations] ([StationPID]) "
