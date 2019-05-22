
"CREATE TABLE [locations] (
	[ID] Counter
	,[Location] Text (255)
	,[Zone] Text (255)
	,[ParksID] Long
	,[type] Text (255)
	,[county] Text (255)
	,[station] Text (255)
	,[ZonePID] Long
	,[StationPID] Long
	,[CountyPID] Long )"

"CREATE INDEX [CountyPID] ON [locations] ([CountyPID]) "

"CREATE INDEX [ParksID] ON [locations] ([CountyPID]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [locations] ([CountyPID])  WITH PRIMARY DISALLOW NULL "

"CREATE INDEX [StationPID] ON [locations] ([CountyPID]) "

"CREATE INDEX [ZonePID] ON [locations] ([CountyPID]) "
