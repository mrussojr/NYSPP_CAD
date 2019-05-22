
"CREATE TABLE [maps] (
	[id] Counter
	,[parkId] Long
	,[mapType] Long
	,[mapName] Text (255)
	,[mapLocation] Text (255) )"

"CREATE INDEX [id] ON [maps] ([mapLocation]) "

"CREATE INDEX [parkId] ON [maps] ([mapLocation]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [maps] ([mapLocation])  WITH PRIMARY DISALLOW NULL "
