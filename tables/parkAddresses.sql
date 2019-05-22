
"CREATE TABLE [parkAddresses] (
	[parksId] Long
	,[address] Text (255)
	,[zipcode] Text (255)
	,[parkName] Text (255) )"

"CREATE INDEX [parksId] ON [parkAddresses] ([parkName]) "

"CREATE INDEX [zipcode] ON [parkAddresses] ([parkName]) "
