
"CREATE TABLE [EventCallAssoc] (
	[ID] Counter
	,[EventID] Long
	,[CallID] Long )"

"CREATE INDEX [CallID] ON [EventCallAssoc] ([CallID]) "

"CREATE INDEX [EventID] ON [EventCallAssoc] ([CallID]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [EventCallAssoc] ([CallID])  WITH PRIMARY DISALLOW NULL "
