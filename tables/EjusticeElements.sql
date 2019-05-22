
"CREATE TABLE [EjusticeElements] (
	[ID] Counter
	,[FormID] Text (255)
	,[ElementType] Text (255)
	,[ElementName] Text (255)
	,[ValueFormName] Text (255)
	,[ValueFieldName] Text (255)
	,[ElementDescrition] Text (255)
	,[OrderId] Long )"

"CREATE INDEX [FormID] ON [EjusticeElements] ([OrderId]) "

"CREATE INDEX [ID] ON [EjusticeElements] ([OrderId]) "

"CREATE INDEX [OrderId] ON [EjusticeElements] ([OrderId]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [EjusticeElements] ([OrderId])  WITH PRIMARY DISALLOW NULL "
