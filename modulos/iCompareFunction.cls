VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iCompareFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function compare(ByRef data As Variant _
                    , ByVal row As Long) As Boolean

    Err.Raise vbObjectError + 601, "iCompareFunction.Filter" _
            , "This function should not be called. You should only the Filter function in a class then Implements this Class."

End Function
