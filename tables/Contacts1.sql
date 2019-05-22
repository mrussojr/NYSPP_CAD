
"CREATE TABLE [Contacts1] (
	[ID] Counter
	,[NAME] Text (255)
	,[SHIELD] Double
	,[RANK] Text (255)
	,[LOCATION] Text (255)
	,[ADDRESS_LINE_1] Text (255)
	,[ADDRESS_LINE_2] Text (255)
	,[PERM] Text (255)
	,[PHONE_1] Text (255)
	,[PHONE_2] Text (255)
	,[PHONE_3] Text (255)
	,[EMAIL] Text (255)
	,[ZONE] Long )"

"CREATE UNIQUE INDEX [ID] ON [Contacts1] ([ZONE])  WITH PRIMARY DISALLOW NULL "
