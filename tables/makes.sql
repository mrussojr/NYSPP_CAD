
"CREATE TABLE [makes] (
	[ID] Counter
	,[makes] Text (255)
	,[FullMakes] Text (255)
	,[ParksID] Long )"

"CREATE INDEX [ID] ON [makes] ([ParksID]) "

"CREATE INDEX [ParksID] ON [makes] ([ParksID]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [makes] ([ParksID])  WITH PRIMARY DISALLOW NULL "
