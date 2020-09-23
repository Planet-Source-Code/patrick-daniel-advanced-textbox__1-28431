VERSION 5.00
Begin VB.UserControl NumericBox 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   ScaleHeight     =   345
   ScaleWidth      =   1590
   ToolboxBitmap   =   "NumericBox.ctx":0000
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "NumericBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'--***********************************************
'-- NOTES FROM ME: PATRICK R. DANIEL
'--***********************************************

'-- This is an extremely chopped down version of the control I acually use.
'-- I have offered this control to PSC to show people how to do true numeric validations
'-- This is only 1 way of doing it...and it is the way I need it to be for my projects.
'-- If anyone has any better ideas, please let me know, I love input.
'--
'-- Email me --> patrick1@mediaone.net
'--
'-- If anyone is wondering, the core Numeric validation code came from www.microsoft.com
'----> CBool(strText = "-" Or strText = "-." Or strText = "." Or IsNumeric(strText))

'-- The Code used to change the Caret when in Insert Mode can from PSS
'-- Author:
'-- Submission:

'-- I have expanded it to take the character sizes into consideration (sorta Unsuccessfully)
'-- You will find that if you change the font size to something other than 8, the caret doesn't
'-- look correct. If anyone wants to expand on it, please send me the code, I am not good with fonts
'-- and don't pretend to be.

'-- This version also allows you to disable Paste
'-- I got that routine from Johan's submission "Disable Paste"

'-- I hope someone gets something out of this..and if so, please write comments and VOTE!

'Default Property Values:
Const m_def_AdvMode As Integer = 1
Const m_def_AdvAllowNegative As Boolean = False
Const m_def_AdvDecimalPlaces As Integer = 0
Const m_def_AdvTabOnEnter As Boolean = False
Const m_def_AdvUCase As Boolean = False
Const m_def_AdvFollowInsertMode As Integer = 1
Const m_def_AdvSelectOnFocus As Integer = 1
Const m_def_AdvNoApostrophe As Boolean = False
Const m_def_NoApostPrompt As Boolean = False
Const m_def_AdvCurrency As Boolean = False
Const m_def_AdvMaxLength As Integer = 0
'Const m_def_AdvDisablePaste As Boolean = False

'Property Variables:
Dim m_AdvMode As Integer
Dim m_AdvAllowNegative As Boolean
Dim m_AdvDecimalPlaces As Integer
Dim m_AdvTabOnEnter As Boolean
Dim m_AdvUCase As Boolean
Dim m_AdvFollowInsertMode As Boolean
Dim m_AdvSelectOnFocus As Boolean
Dim m_AdvNoApostrophe As Boolean
Dim m_AdvNoApostPrompt As Boolean
Dim m_AdvCurrency As Boolean
Dim m_AdvMaxLength As Integer
'Dim m_AdvDisablePaste As Boolean

Dim mblnInsertMode As Boolean
Dim mblnHasFocus As Boolean

Enum ModeConstants
   NumericMode = 0
   TextMode
End Enum

Enum AppearanceConstants
   advFlat = 0
   adv3D
End Enum

Enum BorderConstants
   advNone = 0
   advSolid
End Enum

Dim mblnHooked As Boolean

'Event Declarations:
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Change()
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   Text1.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

Public Property Get DataChanged() As Boolean
Attribute DataChanged.VB_MemberFlags = "400"
   DataChanged = Text1.DataChanged
End Property

Public Property Let DataChanged(blnChanged As Boolean)
   Text1.DataChanged = blnChanged
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
   ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   Text1.ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
   Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled = New_Enabled
   Text1.Enabled = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
   Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set Text1.Font = New_Font
   PropertyChanged "Font"
End Property

Public Property Get BorderStyle() As BorderConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
   BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderConstants)
   Text1.BorderStyle = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
   Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
   Text1.Alignment() = New_Alignment
   PropertyChanged "Alignment"
End Property

Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
   Appearance = Text1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
   Text1.Appearance() = New_Appearance
   PropertyChanged "Appearance"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
   Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
   Text1.Locked() = New_Locked
   PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
   MaxLength = m_AdvMaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
   
   m_AdvMaxLength = New_MaxLength
   
   If m_AdvMode = 1 Then
      Text1.MaxLength = New_MaxLength
   End If
   
   PropertyChanged "MaxLength"
   
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
   Set MouseIcon = Text1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
   Set Text1.MouseIcon = New_MouseIcon
   PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
   MousePointer = Text1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
   Text1.MousePointer() = New_MousePointer
   PropertyChanged "MousePointer"
End Property

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   
   '-- If the user has presses the insert key we need to validate
   '-- Also, if we are in insert mode we must validate every move
   '-- So the caret gets set the the appropriate width for the character we are on
   If KeyCode = vbKeyInsert Or mblnInsertMode Then
      
      ValidateInsertMode
      
   End If
   
   RaiseEvent KeyUp(KeyCode, Shift)
   
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Public Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
   OLEDragMode = Text1.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
   Text1.OLEDragMode() = New_OLEDragMode
   PropertyChanged "OLEDragMode"
End Property

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
   RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
   OLEDropMode = Text1.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
   Text1.OLEDropMode() = New_OLEDropMode
   PropertyChanged "OLEDropMode"
End Property

Public Property Get PasswordChar() As String
   PasswordChar = Text1.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
   Text1.PasswordChar() = New_PasswordChar
   PropertyChanged "PasswordChar"
End Property

Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
   ScrollBars = Text1.ScrollBars
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
   SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal lSelLength As Long)
   Text1.SelLength() = lSelLength
   PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
Attribute SelStart.VB_MemberFlags = "400"
   SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal lSelStart As Long)
   Text1.SelStart() = lSelStart
   PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
   SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal strSelText As String)
   Text1.SelText() = strSelText
   PropertyChanged "SelText"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "1034"
On Error Resume Next
   
   '-- If we are in currency mode, we need to convert this number
   '-- Back to a regular number because we are probably retrieving it
   '-- to store in a database..I use CDbl to convert it to a double
   
   If AdvMode = NumericMode And m_AdvCurrency Then
   
      Text = CStr(CDbl(Text1.Text))
   
   Else
      
      Text = Text1.Text
      
   End If
   
End Property

Public Property Let Text(ByVal New_Text As String)
On Error Resume Next

   '-- WARNING - I do not validate this text. Text validations are ONLY made as the user types
   '-- I assume all text put into the textbox through this event has already been validated
   '-- and is correct. This control is primarily designed for database use.  No invalid data should
   '-- be saved to the database...thus avoiding garbage in / garbage out...and eliminating the need for
   '-- redundant validations both in and out.
   
   '-- As of now, the only way a person can hit a Garbage in / out scenerio is if the control is set to
   '-- numeric and the user pastes alpha text into it.  Validations will not catch it.  Surprisingly, I
   '-- have never ran into this scenerio. User don't usually do this...or they do and I have just gotten
   '-- lucky these past few years :-)
   
   If m_AdvUCase Then
      
      New_Text = UCase(New_Text)
   
   End If
   
   If UserControl.Extender.Parent.ActiveControl.Name = UserControl.Extender.Name Then
      
      If m_AdvMode = 0 Then
         
         If m_AdvMaxLength > 0 Then
            
            Text1.Text = Left(New_Text, m_AdvMaxLength)
         
         Else
            
            Text1.Text = New_Text
            
         End If
      
      Else
         
         Text1.Text = New_Text
      
      End If
      
   Else
      
      If m_AdvCurrency And m_AdvMode = NumericMode Then
         
         If m_AdvMaxLength > 0 Then
         
            Text1.Text = Left(Format(New_Text, "$0.00"), m_AdvMaxLength)
         
         Else
         
            Text1.Text = Format(New_Text, "$0.00")
         
         End If
         
      Else
         
         If m_AdvMaxLength > 0 Then
         
            Text1.Text = Left(New_Text, m_AdvMaxLength)
         
         Else
            
            Text1.Text = New_Text
         
         End If
      
      End If
      
   End If
   
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
   ToolTipText = Text1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
   Text1.ToolTipText() = New_ToolTipText
   PropertyChanged "ToolTipText"
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
   WhatsThisHelpID = Text1.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
   Text1.WhatsThisHelpID() = New_WhatsThisHelpID
   PropertyChanged "WhatsThisHelpID"
End Property

Public Property Get AdvCurrencyFormat() As Boolean
   AdvCurrencyFormat = m_AdvCurrency
End Property

Public Property Let AdvCurrencyFormat(ByVal blnTemp As Boolean)
   m_AdvCurrency = blnTemp
   PropertyChanged "CurrencyFormat"
End Property

Public Property Get AdvMode() As ModeConstants
   AdvMode = m_AdvMode
End Property

Public Property Let AdvMode(ByVal New_AdvMode As ModeConstants)
   m_AdvMode = New_AdvMode
   PropertyChanged "AdvMode"
End Property

Public Property Get AdvAllowNegative() As Boolean
   AdvAllowNegative = m_AdvAllowNegative
End Property

Public Property Let AdvAllowNegative(ByVal New_AdvAllowNegative As Boolean)
   m_AdvAllowNegative = New_AdvAllowNegative
   PropertyChanged "AdvAllowNegative"
End Property

Public Property Let AdvNoApostrophe(ByVal new_Value As Boolean)
   m_AdvNoApostrophe = new_Value
   PropertyChanged "AdvNoApostrophe"
End Property

Public Property Get AdvNoApostrophe() As Boolean
   AdvNoApostrophe = m_AdvNoApostrophe
End Property

Public Property Let AdvNoApostPrompt(new_Value As Boolean)
   m_AdvNoApostPrompt = new_Value
   PropertyChanged "AdvNoApostPrompt"
End Property

Public Property Get AdvNoApostPrompt() As Boolean
   AdvNoApostPrompt = m_AdvNoApostPrompt
End Property

Public Property Get AdvDecimalPlaces() As Integer
   AdvDecimalPlaces = m_AdvDecimalPlaces
End Property

Public Property Let AdvDecimalPlaces(ByVal New_AdvDecimalPlaces As Integer)
   m_AdvDecimalPlaces = New_AdvDecimalPlaces
   PropertyChanged "AdvDecimalPlaces"
End Property

Public Property Get AdvTabOnEnter() As Boolean
   AdvTabOnEnter = m_AdvTabOnEnter
End Property

Public Property Let AdvTabOnEnter(ByVal New_AdvTabOnEnter As Boolean)
   m_AdvTabOnEnter = New_AdvTabOnEnter
   PropertyChanged "AdvTabOnEnter"
End Property

Public Property Get AdvUCase() As Boolean
   AdvUCase = m_AdvUCase
End Property

Public Property Let AdvUCase(ByVal New_AdvUCase As Boolean)
   m_AdvUCase = New_AdvUCase
   PropertyChanged "AdvUCase"
End Property

Public Property Get AdvFollowInsertMode() As Boolean
   AdvFollowInsertMode = m_AdvFollowInsertMode
End Property

Public Property Let AdvFollowInsertMode(ByVal New_AdvFollowInsertMode As Boolean)
   m_AdvFollowInsertMode = New_AdvFollowInsertMode
   PropertyChanged "AdvFollowInsertMode"
End Property

Public Property Get AdvSelectOnFocus() As Boolean
   AdvSelectOnFocus = m_AdvSelectOnFocus
End Property

Public Property Let AdvSelectOnFocus(ByVal New_AdvSelectOnFocus As Boolean)
   m_AdvSelectOnFocus = New_AdvSelectOnFocus
   PropertyChanged "AdvSelectOnFocus"
End Property

'Public Property Get AdvDisablePaste() As Boolean
'   AdvDisablePaste = m_AdvDisablePaste
'End Property
'
'Public Property Let AdvDisablePaste(ByVal blnValue As Boolean)
'   m_AdvDisablePaste = blnValue
'   PropertyChanged "DisablePaste"
'End Property

Private Sub Text1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
   RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub Text1_OLESetData(Data As DataObject, DataFormat As Integer)
   RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Text1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
   RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_EnterFocus()
On Error Resume Next
   
   If mblnHasFocus Then Exit Sub
   
   mblnHasFocus = True
   
   With Text1
      
      If AdvMode = NumericMode And m_AdvCurrency Then
      
         .Text = Format(.Text, "0.00")
      
      End If
      
      If AdvSelectOnFocus Then
   
         .SelStart = 0
         .SelLength = Len(.Text)

      End If
      
      If .Enabled Then
         
         .SetFocus
         
      End If
      
   End With
   
   '-- Since we have entered..we must validate our current insert mode
   ValidateInsertMode
   
End Sub

Private Sub UserControl_ExitFocus()
On Error Resume Next
   
   mblnHasFocus = False
   
   If AdvMode = NumericMode And m_AdvCurrency Then
      
      With Text1
         
         .Text = Format(.Text, "$0.00")
      
      End With
      
   End If
   
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

   Text1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
   
   If UserControl.Height < Text1.Height Then
      UserControl.Height = Text1.Height
   End If
   
End Sub

Public Sub Refresh()
   Text1.Refresh
End Sub

Private Sub Text1_Click()
   RaiseEvent Click
End Sub

Private Sub Text1_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub Text1_OLECompleteDrag(Effect As Long)
   RaiseEvent OLECompleteDrag(Effect)
End Sub

Public Sub OLEDrag()
   Text1.OLEDrag
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   
   Text1.MaxLength = 0
   
   m_AdvMode = m_def_AdvMode
   m_AdvAllowNegative = m_def_AdvAllowNegative
   m_AdvDecimalPlaces = m_def_AdvDecimalPlaces
   m_AdvTabOnEnter = m_def_AdvTabOnEnter
   m_AdvUCase = m_def_AdvUCase
   m_AdvSelectOnFocus = m_def_AdvSelectOnFocus
   m_AdvFollowInsertMode = m_def_AdvFollowInsertMode
   m_AdvCurrency = m_def_AdvCurrency
   
   Font = UserControl.Parent.Font.Name
   
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
   Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   Text1.Enabled = PropBag.ReadProperty("Enabled", True)
   Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
   Text1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
   Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
   Text1.Appearance = PropBag.ReadProperty("Appearance", 1)
   Text1.Locked = PropBag.ReadProperty("Locked", False)
   m_AdvMaxLength = PropBag.ReadProperty("MaxLength", 0)
   Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
   Text1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
   Text1.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
   Text1.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
   Text1.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
   Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
   Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
   Text1.SelText = PropBag.ReadProperty("SelText", "")
   Text1.Text = PropBag.ReadProperty("Text", "AdvTextBox")
   Text1.Tag = PropBag.ReadProperty("Text", "AdvTextBox")
   Text1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
   Text1.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
   m_AdvMode = PropBag.ReadProperty("AdvMode", m_def_AdvMode)
   m_AdvAllowNegative = PropBag.ReadProperty("AdvAllowNegative", m_def_AdvAllowNegative)
   m_AdvDecimalPlaces = PropBag.ReadProperty("AdvDecimalPlaces", m_def_AdvDecimalPlaces)
   m_AdvTabOnEnter = PropBag.ReadProperty("AdvTabOnEnter", m_def_AdvTabOnEnter)
   m_AdvUCase = PropBag.ReadProperty("AdvUCase", m_def_AdvUCase)
   m_AdvFollowInsertMode = PropBag.ReadProperty("AdvFollowInsertMode", m_def_AdvFollowInsertMode)
   m_AdvSelectOnFocus = PropBag.ReadProperty("AdvSelectOnFocus", m_def_AdvSelectOnFocus)
   m_AdvNoApostrophe = PropBag.ReadProperty("AdvNoApostrophe", m_def_AdvNoApostrophe)
   m_AdvNoApostPrompt = PropBag.ReadProperty("AdvNoApostPrompt", m_def_NoApostPrompt)
   m_AdvCurrency = PropBag.ReadProperty("CurrencyFormat", m_def_AdvCurrency)
   'm_AdvDisablePaste = PropBag.ReadProperty("DisablePaste", m_def_AdvDisablePaste)
   
   '-- Set the maxlength in textbox only if in textmode
   '-- Number mode is handled in code
   If m_AdvMode = 1 Then
      
      Text1.MaxLength = m_AdvMaxLength
   
   End If
   
   UserControl.Extender.DataChanged = False
   Text1.DataChanged = False
   
End Sub

'Private Sub UserControl_Show()
'On Error Resume Next
'
'   If Not AdvDisablePaste Then Exit Sub
'
'   If Ambient.UserMode And Not mblnHooked Then
'
'      Hook Text1
'
'      mblnHooked = True
'
'   End If
'
'End Sub
'
'Private Sub UserControl_Terminate()
'On Error Resume Next
'
'   If mblnHooked Then
'
'      UnHook Text1
'
'      mblnHooked = False
'
'   End If
'
'End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
   Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
   Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
   Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
   Call PropBag.WriteProperty("BorderStyle", Text1.BorderStyle, 1)
   Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
   Call PropBag.WriteProperty("Appearance", Text1.Appearance, 1)
   Call PropBag.WriteProperty("Locked", Text1.Locked, False)
   Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
   Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
   Call PropBag.WriteProperty("MousePointer", Text1.MousePointer, 0)
   Call PropBag.WriteProperty("OLEDragMode", Text1.OLEDragMode, 0)
   Call PropBag.WriteProperty("OLEDropMode", Text1.OLEDropMode, 0)
   Call PropBag.WriteProperty("PasswordChar", Text1.PasswordChar, "")
   Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
   Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
   Call PropBag.WriteProperty("SelText", Text1.SelText, "")
   Call PropBag.WriteProperty("Text", Text1.Text, "AdvTextBox")
   Call PropBag.WriteProperty("ToolTipText", Text1.ToolTipText, "")
   Call PropBag.WriteProperty("WhatsThisHelpID", Text1.WhatsThisHelpID, 0)
   Call PropBag.WriteProperty("AdvMode", m_AdvMode, m_def_AdvMode)
   Call PropBag.WriteProperty("AdvAllowNegative", m_AdvAllowNegative, m_def_AdvAllowNegative)
   Call PropBag.WriteProperty("AdvDecimalPlaces", m_AdvDecimalPlaces, m_def_AdvDecimalPlaces)
   Call PropBag.WriteProperty("AdvTabOnEnter", m_AdvTabOnEnter, m_def_AdvTabOnEnter)
   Call PropBag.WriteProperty("AdvUCase", m_AdvUCase, m_def_AdvUCase)
   Call PropBag.WriteProperty("AdvFollowInsertMode", m_AdvFollowInsertMode, m_def_AdvFollowInsertMode)
   Call PropBag.WriteProperty("AdvSelectOnFocus", m_AdvSelectOnFocus, m_def_AdvSelectOnFocus)
   Call PropBag.WriteProperty("AdvNoApostrophe", m_AdvNoApostrophe, m_def_AdvNoApostrophe)
   Call PropBag.WriteProperty("AdvNoApostPrompt", m_AdvNoApostPrompt, m_def_NoApostPrompt)
   Call PropBag.WriteProperty("CurrencyFormat", m_AdvCurrency, m_def_AdvCurrency)
   'Call PropBag.WriteProperty("DisablePaste", m_AdvDisablePaste, m_def_AdvDisablePaste)
   
End Sub

Private Sub Text1_Change()
On Error Resume Next

   RaiseEvent Change
   
   UserControl.Extender.DataChanged = True
   
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next

   '-- If we are insert mode we have to determine which direction to select
   '-- text...We select in reverse if the user is hitting the backspace key
   If (m_AdvFollowInsertMode And mblnInsertMode) And Text1.SelLength = 0 Then
      
      If KeyAscii = 8 Then
         Text1.SelStart = IIf(Text1.SelStart <> 0, Text1.SelStart - 1, 0)
      End If
      
      Text1.SelLength = 1
      
   End If
   
   ValidateText KeyAscii
   
   If KeyAscii = 0 And (m_AdvFollowInsertMode And mblnInsertMode) And SelLength = 1 Then
      Text1.SelLength = 0
   End If
   
End Sub

Private Function ValidateText(KeyAscii As Integer) As Boolean
On Error Resume Next

   Dim intDecimalPosition As Integer
   Dim strText As String
   
   RaiseEvent KeyPress(KeyAscii)
   
   '---------- If Enter Key Pressed Tab On
   If KeyAscii = 13 Then

      If AdvTabOnEnter Then

         KeyAscii = 0
         SendKeys "{TAB}"
         Exit Function

      End If

   End If

   '---------- If ESC Key Pressed Undo Edit
   If KeyAscii = 27 Then

      Call SendMessage(Text1.hwnd, &H304, &O0, &O0)
      KeyAscii = 0
      Exit Function

   End If
   
   If AdvMode = NumericMode Then
      
      '-- Special Characters?
      If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
   
         KeyAscii = 0
         Exit Function
      
      '-- Is this a Control Key? If so exit validations
      ElseIf KeyAscii = 8 Or (KeyAscii > 32 And KeyAscii < 45) Then
         
         Exit Function
      
      '-- Are we allowing Negative Numbers? If not, fail validation
      ElseIf KeyAscii = 45 And Not m_AdvAllowNegative Then
         
         KeyAscii = 0
         Exit Function
         
      '-- Are we allowing decimals? If not, fail validation
      ElseIf KeyAscii = 46 And m_AdvDecimalPlaces = 0 Then
      
         KeyAscii = 0
         Exit Function
         
      End If
      
      '-- Validate the string to see if it is truly numeric
      
      strText = Text1.Text
      strText = Mid(strText, 1, Text1.SelStart) & Chr(KeyAscii) & Mid(strText, Text1.SelStart + Text1.SelLength + 1)
      
      '-- Do we have a string?
      If strText = "" Then Exit Function
      
      '-- If it is not Numeric then fail it
      If Not CBool(strText = "-" Or strText = "-." Or strText = "." Or IsNumeric(strText)) Then
      
         KeyAscii = 0
         Exit Function
      
      End If
      
      '-- Do we already have a decimal?
      If InStr(1, strText, ".") Then
         
         '-- Is the number we entered after the decimal position
         If Text1.SelStart >= InStr(1, strText, ".") Then
            
            '-- Is this number going to exceed the allowable decimal places?
            If Len(Mid(strText, InStr(1, strText, ".") + 1)) > m_AdvDecimalPlaces Then
               
               KeyAscii = 0
               Exit Function
            
            End If
         
         End If
         
      End If
      
      '-- Did user enter a Decimal?
      If KeyAscii = 46 Then
         
         '-- Is this decimal within the allowable decimal places?
         If Len(Mid(strText, InStr(1, strText, ".") + 1)) > m_AdvDecimalPlaces Then
            
            KeyAscii = 0
            Exit Function
         
         End If
      
      End If
      
      '-- Now we must validate maxlength
      If m_AdvCurrency Then
         
         If Len(Format(strText, "0.00")) > m_AdvMaxLength And m_AdvMaxLength > 0 Then
            KeyAscii = 0
            Exit Function
         End If
   
      Else
         
         If Len(strText) > m_AdvMaxLength And m_AdvMaxLength > 0 Then
            
            KeyAscii = 0
            Exit Function
         
         End If
         
      End If
      
   Else
      
      '-- Check for Insert Mode
      If KeyAscii <> 8 Then
         
         If KeyAscii = 39 And AdvNoApostrophe Then   '-- Check For Apostrophe
            
            If AdvNoApostPrompt Then
               
               MsgBox "The Apostrophe(') character is not Allowed for this Field", vbInformation, "Validation"
               
               UserControl.Extender.SetFocus
            
            End If
            
            KeyAscii = 0
            
            Exit Function
         
         End If
         
         If AdvUCase Then
            
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
         
         End If
      
      End If
   
   End If
   
End Function

Private Sub ValidateInsertMode()
On Error Resume Next
   
   If Not m_AdvFollowInsertMode Then Exit Sub
   
   Dim tm As TEXTMETRIC
   Dim dc As Long
   Dim lWidth As Long
   Dim lret As Long
   Dim strChar As String
   
   dc = GetDC(Text1.hwnd)
   
   GetTextMetrics dc, tm
   
   strChar = Mid(Text1.Text, Text1.SelStart + 1, 1)
   
   If strChar = "" Then
      lWidth = tm.tmAveCharWidth
   Else
      lret = GetCharWidth32(dc, Asc(strChar), Asc(strChar), lWidth)
   End If
   
   ReleaseDC Text1.hwnd, dc
   
   If GetKeyState(45) = 1 Then
      
      mblnInsertMode = True
      
      CreateCaret Text1.hwnd, 0, lWidth - 1, tm.tmAscent
      ShowCaret Text1.hwnd
      
   Else
         
      mblnInsertMode = False
         
      CreateCaret Text1.hwnd, 0, 1, tm.tmAscent
      ShowCaret Text1.hwnd
   
   End If
   
   Text1.SetFocus
   
End Sub
