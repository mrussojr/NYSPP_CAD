
"CREATE TABLE [people] (
	[ID] Counter
	,[LastName] Text (255)
	,[FirstName] Text (255)
	,[Middle] Text (255)
	,[DOB] Text (255)
	,[Sex] Text (255)
	,[State] Text (255)
	,[CID] Text (255)
	,[Class] Text (255)
	,[Expiration] Text (255)
	,[Height] Text (255)
	,[EyeColor] Text (255)
	,[StreetAddress] Text (255)
	,[County] Text (255)
	,[Municipality] Text (255)
	,[ZipCode] Long
	,[Event] Long )"

"CREATE INDEX [CID] ON [people] ([Event]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [people] ([Event])  WITH PRIMARY DISALLOW NULL "

"CREATE INDEX [ZipCode] ON [people] ([Event]) "
