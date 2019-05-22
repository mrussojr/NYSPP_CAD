Attribute VB_Name = "Form_ActiveEventsForEJustice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command19_Click()
    'Set the eventID variable of the ejustice form to the eventID of the event selected
    Forms.Item("EJustice").EventID = Me.idBox.Value
    'close the form
    DoCmd.Close acForm, "ActiveEventsForEJustice"
    'run the populate table function from the the Ejustice form
    Forms.Item("EJustice").populateTable
End Sub
