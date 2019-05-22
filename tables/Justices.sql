
"CREATE TABLE [Justices] (
	[LastName] Text (50)
	,[FirstName] Text (50)
	,[HomePhone] Text (30)
	,[Cell/Pager/Work] Text (30)
	,[OtherOffice] Text (50)
	,[Day] Text (50)
	,[Time] Text (50)
	,[Notes] Text (255)
	,[EMAIL] Text (50)
	,[CourtID] Long
	,[ID] Counter )"

"CREATE INDEX [ID] ON [Justices] ([ID]) "

"CREATE INDEX [ID1] ON [Justices] ([ID]) "

"CREATE INDEX [LastName] ON [Justices] ([ID]) "
