
"CREATE TABLE [ReasonCodes] (
	[ID] Counter
	,[ReasonCode] Text (255) )"

"CREATE UNIQUE INDEX [PrimaryKey] ON [ReasonCodes] ([ReasonCode])  WITH PRIMARY DISALLOW NULL "

"CREATE INDEX [ReasonCode] ON [ReasonCodes] ([ReasonCode]) "
