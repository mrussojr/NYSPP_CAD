
"CREATE TABLE [commlog] (
	[Date1] Text (255)
	,[Time1] Text (255)
	,[UnitCalled] Text (255)
	,[SourceCall] Text (255)
	,[Reason] Text (255)
	,[Narrative] Memo
	,[Dispatcher] Text (255)
	,[id] Counter
	,[Event] Text (255) )"

"CREATE INDEX [id] ON [commlog] ([Event]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [commlog] ([Event])  WITH PRIMARY DISALLOW NULL "
