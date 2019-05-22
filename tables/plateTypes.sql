
"CREATE TABLE [plateTypes] (
	[ID] Counter
	,[Type] Text (255)
	,[TypeNo] Long
	,[ShortCode] Text (255) )"

"CREATE UNIQUE INDEX [PrimaryKey] ON [plateTypes] ([ShortCode])  WITH PRIMARY DISALLOW NULL "

"CREATE INDEX [ShortCode] ON [plateTypes] ([ShortCode]) "
