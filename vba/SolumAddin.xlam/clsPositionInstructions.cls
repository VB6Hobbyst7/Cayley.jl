VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPositionInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsPositionInstructions
' Author    : Philip Swannell
' Date      : 15-Nov-2013
' Purpose   : Trivial class module for use as argument to ShowCommandBarPopup or SetFormPosition
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Public AnchorObject As Object
Public X_Nudge As Double        '<-- measured in Pixels when the class is used as an argument to ShowCommandBarPopup, _
                                 measured in Points when used as an an argument to SetFormPosition _
                                 function fX() can be used to convert Points to pixels
Public Y_Nudge As Double        '<-- ditto ...fY()...
