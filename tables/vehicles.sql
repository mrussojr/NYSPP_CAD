
"CREATE TABLE [vehicles] (
	[ID] Counter
	,[State] Text (255)
	,[PlateNumber] Text (255)
	,[Type] Long
	,[Expiration] Text (255)
	,[VehYear] Long
	,[VehMake] Text (255)
	,[VIN] Text (255)
	,[Status] Text (255)
	,[PlateIssued] Text (255)
	,[PlateStyle] Text (255)
	,[InsuranceCode] Long
	,[VehStyle] Text (255)
	,[VehModel] Text (255)
	,[VehColor] Text (255)
	,[Event] Long )"

"CREATE INDEX [InsuranceCode] ON [vehicles] ([Event]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [vehicles] ([Event])  WITH PRIMARY DISALLOW NULL "
