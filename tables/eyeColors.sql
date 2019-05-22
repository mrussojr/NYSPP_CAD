
"CREATE TABLE [eyeColors] (
	[ID] Counter
	,[Color] Text (255)
	,[ParksId] Long )"

"CREATE INDEX [ParksId] ON [eyeColors] ([ParksId]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [eyeColors] ([ParksId])  WITH PRIMARY DISALLOW NULL "
