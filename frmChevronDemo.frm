VERSION 5.00
Object = "{56034C77-7369-44EE-828E-D0C208D19BDF}#8.0#0"; "Gold Button.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChevronDemo 
   Caption         =   "Chevron Demo by Joseph M. Ferris"
   ClientHeight    =   4800
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDummy 
      BackColor       =   &H80000005&
      Height          =   4275
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   5610
      TabIndex        =   3
      Top             =   525
      Width           =   5670
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   4875
      Top             =   4005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChevronDemo.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChevronDemo.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChevronDemo.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChevronDemo.frx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChevronDemo.frx":0448
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChevronDemo.frx":055A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraToolbarContainer 
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   5670
      Begin MSComctlLib.Toolbar tlbMainToolbar 
         Height          =   390
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   688
         ButtonWidth     =   609
         Wrappable       =   0   'False
         ImageList       =   "imlToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Open"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Null"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Save"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Print"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Null"
               ImageIndex      =   6
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Undo"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "Redo"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin pOCX.GoldButton cmdGoldenChevron 
         Height          =   435
         Left            =   5445
         TabIndex        =   1
         Top             =   105
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   767
         Caption         =   ">>"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnHover         =   5
      End
   End
   Begin VB.Menu mnuPopUpParent 
      Caption         =   "PopUpParent"
      Visible         =   0   'False
      Begin VB.Menu mnuChildItem 
         Caption         =   "Child Item"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmChevronDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*' Chevrons in Toolbars
'*'
'*' Author  : Joseph M. Ferris  (joseph.ferris@cdicorp.com)
'*'                             (jferris@jdfdesign.com)
'*'
'*' Date    : August 29, 2000
'*'
'*' Purpose : To implement Chevrons in a toolbar to display items that would
'*'           normally "dissappear" when a form is sized too small via a
'*'           dynamic popup menu.
'*'
'*' Special Thanks:
'*'
'*'     A very special thank you to Night Wolf for the incredible and lightweight
'*'     Gold Button control, which is used in this project.
'*'
'*' This source code is in the public domain.  Feel free to use this source
'*' code however you wish, as long as credit is given to the original author.

Public intFromChevron As Integer        '*' Public Integer to store user selection

Private Sub Form_Load()

'*' The first thing to do when the form loads is identify which items
'*' on the toolbar will have an entry on the popup menu.  A value has
'*' been added to the Tag property of each individual button.  This can
'*' be viewed by looking at the Property Sheet for the toolbar.
'*'
'*' A value of "Null" tells our loop contructor to ignore this item
'*' as a button.  We use this to exclude item seperators and other styles
'*' that are possible for a button.
'*'
'*' Any value other than null is used for the description in the popupmenu.

Dim intConst            '*' Integer for loop construction
Dim intCounter          '*' Integer for item counter

'*' We will check to see what the next available Index is on the mnuChildItem.

intCounter = mnuChildItem(mnuChildItem.UBound).Index + 1

'*' Next, loop through all of the buttons to determine their status in regards
'*' to being added to the menu.

For intConst = 1 To tlbMainToolbar.Buttons.Count

'*' Since the only menu items that will be added will not contain "Null" as
'*' their Tag, we will search for them now.

If tlbMainToolbar.Buttons(intConst).Tag <> "Null" Then
     
    Load mnuChildItem(intCounter)    '*' We will now load an instance of the blank
                                     '*' menu item.  Since we have encountered a
                                     '*' non-"Null" value, it is valid for our
                                     '*' popup menu.
    
    '*' We will need to set that Tag value as the caption for the menu item.
    '*' Note that we use the intCounter variable to assign the Index to the
    '*' menu item and not the Index of the button.
    
    mnuChildItem(intCounter).Caption = tlbMainToolbar.Buttons(intConst).Tag
    
    mnuChildItem(intCounter).Visible = True     '*' Ensure visibility
    
    intCounter = intCounter + 1                 '*' Increment counter
    
End If

Next intConst                           '*' Continue loop iteration

mnuChildItem(0).Visible = False         '*' The first item in the mnuChildItem
                                        '*' Collection is visible.  Since we have
                                        '*' added items to the Collection, we
                                        '*' can disable the first.  To see how
                                        '*' this is set up, look at the Menu Editor
                                        '*' to see the menu structure that must
                                        '*' exist before running this.
                                        '*'
                                        '*' At least one child item needs to be
                                        '*' visible in a menu with subitems, or
                                        '*' an error will be displayed.

End Sub


Private Sub Form_Resize()

On Error Resume Next

fraToolbarContainer.Width = Me.Width - 110
cmdGoldenChevron.Left = fraToolbarContainer.Width + fraToolbarContainer.Left - cmdGoldenChevron.Width - 10
tlbMainToolbar.Width = fraToolbarContainer.Width + fraToolbarContainer.Left - cmdGoldenChevron.Width - 130
picDummy.Width = fraToolbarContainer.Width

Call CheckVisibility                '*' CheckVisibility is called whenever the
                                    '*' form is resized to determine if the button
                                    '*' is still visible or not.

End Sub

Private Sub cmdGoldenChevron_Click()

'*' When the GoldenChevron is click, we need to display the popup menu.
Me.PopupMenu mnuPopUpParent, , cmdGoldenChevron.Left - 20, cmdGoldenChevron.Top + cmdGoldenChevron.Height - 40

End Sub

Private Sub CheckVisibility()

'*' This subroutine will check each button to see if it is displayed beyond the
'*' the border of the toolbar.

On Error Resume Next                '*' Ignore errors on minimize and restore

Dim intFirstInvisible As Integer    '*' Integer value to store the Index of our first
                                    '*' invisible item.
                                    
Dim intConst As Integer             '*' Integer loop constructor

Dim intCounter As Integer           '*' We need a counter to track the physical location in
intCounter = 1                      '*' the menu Collection.

For intConst = 1 To tlbMainToolbar.Buttons.Count

    '*' Remember, we only work with non-"Null" Tag values.
    
    If tlbMainToolbar.Buttons(intConst).Tag <> "Null" Then
    
        '*' We will now determine if any piece of the current button is further left than
        '*' the toolbar boundary.
        
        If tlbMainToolbar.Buttons(intConst).Left + tlbMainToolbar.Buttons(intConst).Width > tlbMainToolbar.Width Then
            
            '*' When execution reaches this point, we have encountered the first item that is
            '*' not completely visible on the toolbar.  Since none of the other buttons past
            '*' this will be visible, we will terminate our loop here.
                    
            intFirstInvisible = intCounter                  '*' Store first invisible value
            intConst = tlbMainToolbar.Buttons.Count + 1     '*' Force loop termination
            
        End If
        
        '*' Increment our counter by one if we have not found an invisible button, so we
        '*' will move through our Collection.
        
        intCounter = intCounter + 1
        
    End If
    
Next intConst

If intFirstInvisible > 0 Then

    '*' When intFirstInvisible is greater than 0, then we have at least one
    '*' item in our popup menu.  We need to turn on the GoldenChevron button.
    
    cmdGoldenChevron.Visible = True
    
    intCounter = 1              '*' Reset our counter for this constructor

    For intConst = 1 To tlbMainToolbar.Buttons.Count

        If tlbMainToolbar.Buttons(intConst).Tag <> "Null" Then
                
                '*' If the button is visible, we will not show the menu item,
                '*' else we will show the menu item.
                
                If intCounter < intFirstInvisible Then
                    mnuChildItem(intCounter).Visible = False
                Else
                    mnuChildItem(intCounter).Visible = True
                End If
                
                '*' If this part of the condition has been met, increment
                '*' our counter variable.
                
                intCounter = intCounter + 1

        End If

    Next intConst

Else

    '*' If the If statement falls through then all of the buttons are visible.
    '*' We need to make sure that our Chevron button is disabled.
    
    cmdGoldenChevron.Visible = False
    
End If

End Sub

Private Sub mnuChildItem_Click(Index As Integer)

'*' Covert the menu item number to the button number, and pass it to the
'*' HandleEvent subroutine

Dim intConst As Integer

'*' We will try to match the Tag from the button and the caption from the
'*' menu to determine which menu item was selected.

For intConst = 1 To tlbMainToolbar.Buttons.Count

    If tlbMainToolbar.Buttons(intConst).Tag = mnuChildItem(Index).Caption Then
    
        '*' When we have found the magic value, we will assign it to the
        '*' global variable intFromChevron and blow up the loop.
        
        intFromChevron = intConst
        intConst = tlbMainToolbar.Buttons.Count + 1
        
    End If
    
Next intConst

'*' Providing that a value was obtained by the loop, we will pass it to our
'*' common event handler, the HandleEvent subroutine.

If intFromChevron > 0 Then
    HandleEvent (intFromChevron)
End If

End Sub

Private Sub tlbMainToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

'*' We will automatically pass the Index of the button pressed to our HandleEvent
'*' subroutine.  By having a common event handler, we do not need to write code
'*' to handle the menu items and button items individually.

HandleEvent (Button.Index)

End Sub

Private Sub HandleEvent(intEventRaised As Integer)

'*' The common event handler is used so that we do not write the same code twice.
'*' The common event handler is used so that we do not write the same code twice.

'*' Both the button's Index and the menu item's Index are the same if the
'*' correlating item is selected.  We will now use a Select statement to return
'*' feedback to test.

Select Case intEventRaised

    Case 1
        
        MsgBox "New was pressed."
        
    Case 2
    
        MsgBox "Open was pressed."
        
    Case 4
    
        MsgBox "Save was pressed."
        
    Case 5
    
        MsgBox "Print was pressed."
        
    Case 7
    
        MsgBox "Undo was pressed."
    
    Case 8
    
        MsgBox "Redo was pressed."
        
    Case Else
    
        MsgBox "Unidentified error."
        
End Select

'*' We need to reset the intFromChevron global variable.  If we do not, the
'*' same condition will be evaluated the next time this subroutine is called.
'*' Very important!

intFromChevron = 0

End Sub
