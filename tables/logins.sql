
"CREATE TABLE [logins] (
	[ID] Counter
	,[Username] Text (255)
	,[LogInDate] DateTime
	,[ErrorReason] Text (255) )"

"CREATE UNIQUE INDEX [PrimaryKey] ON [logins] ([ErrorReason])  WITH PRIMARY DISALLOW NULL "
