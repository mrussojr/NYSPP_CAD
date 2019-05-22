
"CREATE TABLE [AllCounties] (
	[ID] Counter
	,[County] Text (255)
	,[ZoneID] Long
	,[pID] Long )"

"CREATE INDEX [pID] ON [AllCounties] ([pID]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [AllCounties] ([pID])  WITH PRIMARY DISALLOW NULL "

"CREATE INDEX [ZoneID] ON [AllCounties] ([pID]) "
