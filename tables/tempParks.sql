
"CREATE TABLE [tempParks] (
	[offId] Long
	,[assignmentId] Long
	,[car] Text (255) )"

"CREATE INDEX [assignmentId] ON [tempParks] ([car]) "

"CREATE INDEX [offId] ON [tempParks] ([car]) "
