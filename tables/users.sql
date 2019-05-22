
"CREATE TABLE [users] (
	[ID] Counter
	,[DispatchID] Long
	,[Username] Text (255)
	,[EJusticeUsername] Text (255)
	,[Password] Text (255)
	,[Salt] Text (255)
	,[Active] Long
	,[Authorized] Long
	,[Admin] Long
	,[Email] Text (255)
	,[EJusticeLinkId] Text (255)
	,[ORI] Text (255)
	,[LoggedInEJustice] Long )"

"CREATE INDEX [DispatchID] ON [users] ([LoggedInEJustice]) "

"CREATE INDEX [EJusticeLinkId] ON [users] ([LoggedInEJustice]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [users] ([LoggedInEJustice])  WITH PRIMARY DISALLOW NULL "
