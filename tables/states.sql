
"CREATE TABLE [states] (
	[ID] Counter
	,[State] Text (255)
	,[StateFull] Text (255) )"

"CREATE INDEX [ID] ON [states] ([StateFull]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [states] ([StateFull])  WITH PRIMARY DISALLOW NULL "
