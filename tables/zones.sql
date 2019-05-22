
"CREATE TABLE [zones] (
	[id] Counter
	,[zone] Text (255)
	,[pId] Long )"

"CREATE INDEX [id] ON [zones] ([pId]) "

"CREATE INDEX [pId] ON [zones] ([pId]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [zones] ([pId])  WITH PRIMARY DISALLOW NULL "
