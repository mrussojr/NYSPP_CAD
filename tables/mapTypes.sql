
"CREATE TABLE [mapTypes] (
	[id] Counter
	,[Map_Types] Text (255) )"

"CREATE INDEX [ID] ON [mapTypes] ([Map_Types]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [mapTypes] ([Map_Types])  WITH PRIMARY DISALLOW NULL "
