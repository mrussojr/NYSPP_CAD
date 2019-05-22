
"CREATE TABLE [MosOfYr] (
	[Id] Counter
	,[MonthName] Text (255)
	,[MonthId] Long )"

"CREATE UNIQUE INDEX [Id] ON [MosOfYr] ([MonthId]) "

"CREATE INDEX [MonthId] ON [MosOfYr] ([MonthId]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [MosOfYr] ([MonthId])  WITH PRIMARY DISALLOW NULL "
