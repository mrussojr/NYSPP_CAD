
"CREATE TABLE [DispatcherIDs] (
	[ID] Counter
	,[DispUserName] Text (255)
	,[DispatcherID] Integer
	,[Active] YesNo )"

"CREATE INDEX [DispatcherID] ON [DispatcherIDs] ([Active]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [DispatcherIDs] ([Active])  WITH PRIMARY DISALLOW NULL "
