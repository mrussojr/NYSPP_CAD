
"CREATE TABLE [PANIC_BOOK_-_CEN] (
	[ID] Long
	,[Region] Text (255)
	,[Park] Text (255)
	,[Park_Facility] Text (50)
	,[Park_Phone] Text (50)
	,[County] Text (50)
	,[Copy_of_Address] Text (255)
	,[Address] Memo
	,[Manager] Text (50)
	,[Mgr_Phone] Text (50)
	,[Mgr_Cell] Text (50)
	,[More_Info_(See_Memo)] YesNo
	,[Resides_In_Park] YesNo
	,[Assistant] Text (50)
	,[Asst_Phone] Text (50)
	,[Asst_Cell] Text (50)
	,[Resides_In_Park1] YesNo
	,[FIRE_EMS_AMBULANCE] Text (50)
	,[POISON_CONTROL] Text (50)
	,[NYSP_EN-CON] Text (50)
	,[SHERIFF] Text (50)
	,[HospNum] Memo
	,[Hospital] Memo
	,[Court] Text (50)
	,[Court Phone] Text (50)
	,[Tow_Truck] Text (50)
	,[Tow_Truck_Name] Text (50)
	,[Tow_Truck_2] Text (255)
	,[Tow_Truck_2_Name] Text (255)
	,[Tow_Truck_3] Text (255)
	,[Tow_Truck_3_Name] Text (255)
	,[Animal_Control] Text (50)
	,[Animal_Control_Name] Text (50)
	,[Veterinarian_Name_Number] Text (50)
	,[Locksmith] Text (50)
	,[Locksmith_Name] Text (50)
	,[Locksmith_2] Text (255)
	,[Locksmith_2_Name] Text (255)
	,[Locksmith_3] Text (255)
	,[Locksmith_3_Name] Text (255)
	,[Payphone_Location_Number] Memo
	,[MEMO] Memo
	,[Local_Agency1] Text (255)
	,[Local_Agency1_Num] Text (255)
	,[Local_Agency2] Text (255)
	,[Local_Agency2_Num] Text (255)
	,[Local_NYSP] Text (255)
	,[Local_NYSP Num] Text (255)
	,[POLICE] Memo
	,[Zone_Stations] Memo
	,[Court_Address] Text (255)
	,[Court_Hours] Text (255)
	,[Jusitice1] Text (255)
	,[Justice1H] Text (255)
	,[Justice1C] Text (255)
	,[Justice2] Text (255)
	,[Justice2H] Text (255)
	,[Justice2C] Text (255)
	,[Building_1] Text (255)
	,[2nd_Contact] Text (255)
	,[1st_Contact] Text (255)
	,[Alarm_Phone] Text (255)
	,[Alarm_Company] Text (255)
	,[Alarm_Codes] Text (255)
	,[Building_2] Text (255)
	,[Codes_2] Text (255)
	,[ALARM_SYSTEM] YesNo
	,[Field1] Text (255)
	,[Field2] Text (50)
	,[ParksID] Long
	,[hazmat] Text (255)
	,[Station] Long )"

"CREATE INDEX [Codes_2] ON [PANIC_BOOK_-_CEN] ([Station]) "

"CREATE INDEX [ID] ON [PANIC_BOOK_-_CEN] ([Station]) "

"CREATE INDEX [ParksID] ON [PANIC_BOOK_-_CEN] ([Station]) "
