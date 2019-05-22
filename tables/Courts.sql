
"CREATE TABLE [Courts] (
	[County] Text (50)
	,[Jurisdiction] Text (50)
	,[Court_Phone] Text (30)
	,[Court_Fax] Text (255)
	,[Mailing_Address] Text (100)
	,[Mailing_Address_2] Text (255)
	,[City_State_Zip] Text (255)
	,[Court_Address] Text (50)
	,[Loc_Code] Text (100)
	,[Court_ORI] Text (50)
	,[Notes] Memo
	,[Office_Hours] Text (255)
	,[ID] Counter )"

"CREATE INDEX [ID] ON [Courts] ([ID]) "

"CREATE INDEX [Loc  Code] ON [Courts] ([ID]) "
