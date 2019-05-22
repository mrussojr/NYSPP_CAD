
"CREATE TABLE [EventPersonAssoc] (
	[ID] Counter
	,[EventId] Long
	,[PersonId] Long )"

"CREATE INDEX [EventId] ON [EventPersonAssoc] ([PersonId]) "

"CREATE INDEX [PersonId] ON [EventPersonAssoc] ([PersonId]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [EventPersonAssoc] ([PersonId])  WITH PRIMARY DISALLOW NULL "
