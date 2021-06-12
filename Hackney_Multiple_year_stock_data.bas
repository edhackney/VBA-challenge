Attribute VB_Name = "Module1"
Public Sub Find_Uniques()

ActiveSheet.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("J1"), Unique:=True

End Sub


