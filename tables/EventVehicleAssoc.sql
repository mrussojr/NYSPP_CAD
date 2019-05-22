
"CREATE TABLE [EventVehicleAssoc] (
	[ID] Counter
	,[EventId] Long
	,[VehId] Long )"

"CREATE INDEX [EventId] ON [EventVehicleAssoc] ([VehId]) "

"CREATE INDEX [ID] ON [EventVehicleAssoc] ([VehId]) "

"CREATE UNIQUE INDEX [PrimaryKey] ON [EventVehicleAssoc] ([VehId])  WITH PRIMARY DISALLOW NULL "

"CREATE INDEX [VehId] ON [EventVehicleAssoc] ([VehId]) "
