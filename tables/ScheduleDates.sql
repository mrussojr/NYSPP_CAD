
"CREATE TABLE [ScheduleDates] (
	[ID] Counter
	,[BeginDate] DateTime
	,[EndDate] DateTime
	,[FileName] Text (255) )"

"CREATE INDEX [ID] ON [ScheduleDates] ([FileName]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [ScheduleDates] ([FileName])  WITH PRIMARY DISALLOW NULL "
