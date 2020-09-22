VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDemo 
   Caption         =   "Right Click TreeView or Node To Add/Delete/Change Node Properties"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   " TreeView "
      Height          =   1455
      Left            =   3720
      TabIndex        =   16
      Top             =   4920
      Width           =   4095
      Begin VB.CommandButton cmdSaveTree 
         Caption         =   "Save Current TreeView To PropertyBag"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   3855
      End
      Begin VB.CommandButton cmdLoadTree 
         Caption         =   "Load TreeView From PropertyBag"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   3855
      End
   End
   Begin MSComctlLib.ImageList demoImgList 
      Left            =   7320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "iClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":307C
            Key             =   "iOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":582E
            Key             =   "iLeaf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6240
            Key             =   "iBAS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B1A
            Key             =   "iCLASS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73F4
            Key             =   "iShapes"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7CCE
            Key             =   "iVB"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85A8
            Key             =   "iCube"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E82
            Key             =   "iDiamonds"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":975C
            Key             =   "iDocPlus"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A036
            Key             =   "iText"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A910
            Key             =   "iSelected"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   " Selected Node Properties "
      Height          =   4695
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtSelectedNodeTag 
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Text            =   "txtSelectedNodeTag"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox cmbNodeChecked 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3720
         Width           =   2175
      End
      Begin VB.ComboBox cmbNodeSorted 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtSelNodeKey 
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Text            =   "txtSelNodeKey"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton cmdApplyChanges 
         Caption         =   "Apply Changes To Treeview (NOT Saved)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   4200
         Width           =   3855
      End
      Begin VB.ComboBox cmbSelectedImage 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3000
         Width           =   2175
      End
      Begin VB.ComboBox cmbRegImg 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox txtSelectedNodeText 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   "txtSelectedNodeText"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Node Tag:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   1440
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Node Checked?"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Node Sorted?"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Image:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Node Image:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1590
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3480
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Node Text"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label lblParentNodeKey 
         BackStyle       =   0  'Transparent
         Caption         =   "lblParentNodeKey"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Node Key:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lblParentNodeIndex 
         BackStyle       =   0  'Transparent
         Caption         =   "lblParentNodeIndex"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Node Index:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Node Key:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1425
      End
      Begin VB.Label lblSelNodeIndex 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSelNodeIndex"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Node Index:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1545
      End
   End
   Begin MSComctlLib.TreeView demoTView 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11033
      _Version        =   393217
      Style           =   7
      ImageList       =   "demoImgList"
      Appearance      =   1
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd2Root 
         Caption         =   "Add Node To Root"
      End
      Begin VB.Menu mnuAddChild 
         Caption         =   "Add Child To Selected Node"
      End
      Begin VB.Menu mnuHyph1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveNode 
         Caption         =   "Remove Selected Node"
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ TreeView Load/Save using propertybag demo                                            ++
'++ Ver. 1.0.1 Written by: CptnVic  29 Feb. 2008                                         ++
'++ Ver. 1.0.0 Written by: CptnVic  7 Feb. 2008                                          ++
'++ Copyright: Author retains ownership and his rights to use this code, however, you may++
'++            use the code in your projects as you wish with no restriction provided    ++
'++            that you leave this copyright notice intact.                              ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ The code in this form's module is just to allow you to play around with the treeview ++
'++ and the SaveTree and LoadTree subs in the module.                                    ++
'++ The SaveTree sub doesn't care how you add/modify/delete nodes in the treeview since  ++
'++ it parses the treeview when the SaveTree sub is called.  Therefore, you can use your ++
'++ own methods for adding, etc., nodes to the treeview.                                 ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ BTW:                                                                                 ++
'++ The form code is not the point of this demo - but if you want to test this project   ++
'++ from scratch, simply delete the demoTView.bag, and demoTView.dat files (included in  ++
'++ the zip file) from the application's directory, run the program and experiment as    ++
'++ you wish.                                                                            ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type CountInfo
    iCnt As Integer 'change to long if you plan to add more than 32,767 items over time
End Type

Dim TreeIsDirty As Boolean 'a flag indication changes in the treeview

Private Sub cmdApplyChanges_Click()
    '-----------------------------------------------------------
    '-- apply changes made in property frame to selected node --
    '-----------------------------------------------------------
    If demoTView.SelectedItem Is Nothing Then Exit Sub 'see: demoTView_MouseUp for details
    
    Dim NodX As Node
    
    demoTView.SelectedItem.Key = txtSelNodeKey
    demoTView.SelectedItem.Text = txtSelectedNodeText
    demoTView.SelectedItem.Tag = txtSelectedNodeTag
    demoTView.SelectedItem.Image = cmbRegImg.Text
    demoTView.SelectedItem.SelectedImage = cmbSelectedImage.Text
    demoTView.SelectedItem.Sorted = cmbNodeSorted.Text
    demoTView.SelectedItem.Checked = cmbNodeChecked
    TreeIsDirty = True

End Sub

Private Sub cmdLoadTree_Click()
    'load/reload tree
    LoadTree demoTView 'Load the treeview - use the default filename
    TreeIsDirty = False
End Sub

Private Sub cmdSaveTree_Click()
    SaveTree demoTView 'save the treeview - use the default filename - do not back up
    TreeIsDirty = False
End Sub

Private Sub demoTView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then PopupMenu mnuPopUp
End Sub

Private Sub demoTView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If demoTView.SelectedItem Is Nothing Then
        '-- If you downloaded version 1.0.0 - you would notice that I formerly tested for
        '-- the selected item by generating an error in var = demoTView.SelectedItem.Index
        '-- This was a bad coding practice I used mimicking some MSDN code I saw during my
        '-- research.  Since the demoTView.SelectedItem.Index generated an error if no item
        '-- had been selected, it worked fine... however, the actual error being generated
        '-- was: (91) "Object variable or With block variable not set".  Therefore, the error
        '-- I was using to exit the sub could not exist if no node had been selected - since
        '-- by definition if demoTView.SelectedItem Is Nothing Then demoTView.SelectedItem.Index
        '-- would not make any sense.
        '-- The short version is: testing if demoTView.SelectedItem Is Nothing makes more sense
        '-- and is a better code practice.
        Exit Sub 'don't show properties
    End If
    
    If Button And vbLeftButton Then
        ShowItemProperties
        cmdApplyChanges.Enabled = True
    End If

End Sub

Private Sub Form_Load()
    
    ClearInfo
    LoadKeys
    LoadTree demoTView 'load the treeview passing the treeview name
    TreeIsDirty = False
End Sub
Private Sub ShowItemProperties()
    '--------------------------------------------------------------------------------------------
    '-- This sub displays the selected node properties as they currently exist in the treeview --
    '--------------------------------------------------------------------------------------------
    If demoTView.SelectedItem Is Nothing Then Exit Sub
    
    lblParentNodeIndex.Caption = GetParent(demoTView.SelectedItem.Index)
    lblParentNodeKey.Caption = GetParentKey(demoTView.SelectedItem.Index)
    lblSelNodeIndex.Caption = demoTView.SelectedItem.Index
    txtSelNodeKey.Text = demoTView.SelectedItem.Key
    txtSelectedNodeText.Text = demoTView.SelectedItem.Text
    txtSelectedNodeTag.Text = demoTView.SelectedItem.Tag
    cmbRegImg.ListIndex = GetIconNum(demoTView.SelectedItem.Image)
    cmbSelectedImage.ListIndex = GetIconNum(demoTView.SelectedItem.SelectedImage)
    cmbNodeSorted.ListIndex = Abs(CInt(demoTView.SelectedItem.Sorted)) 'convert boolean to a useful listindex value
    cmbNodeChecked.ListIndex = Abs(CInt(demoTView.SelectedItem.Checked)) 'ditto
    
End Sub

Private Sub ClearInfo()
    '---------------------------------------------------
    '-- This sub just clears the selected node labels --
    '---------------------------------------------------
    lblParentNodeIndex.Caption = ""
    lblParentNodeKey.Caption = ""
    lblSelNodeIndex.Caption = ""
    txtSelNodeKey.Text = ""
    txtSelectedNodeText.Text = ""
    txtSelectedNodeTag.Text = ""
    cmbRegImg.ListIndex = -1
    cmbSelectedImage.ListIndex = -1
    cmbNodeSorted.ListIndex = -1
    cmbNodeChecked.ListIndex = -1
End Sub
Private Sub LoadKeys()
    '-------------------------------------------------------------------
    '-- Load available icons to combo boxes and sorted/checked values --
    '-------------------------------------------------------------------
    Dim i As Integer
    
    cmbRegImg.Clear
    cmbSelectedImage.Clear
    For i = 1 To demoImgList.ListImages.Count
        cmbRegImg.AddItem demoImgList.ListImages(i).Key
        cmbSelectedImage.AddItem demoImgList.ListImages(i).Key
    Next
    cmbNodeSorted.Clear
    cmbNodeSorted.AddItem "False"
    cmbNodeSorted.AddItem "True"
    cmbNodeChecked.Clear
    cmbNodeChecked.AddItem "False"
    cmbNodeChecked.AddItem "True"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '-------------------------------------
    '-- Check treeview file saved state --
    '-------------------------------------
    Dim Response As Integer
    If TreeIsDirty Then
        Response = MsgBox("Do You Want To Save The TreeView Contents Before Exiting?", vbYesNo + vbQuestion, "TreeView Contents Have Not Been Saved!")
        If Response = vbYes Then 'save the treeview before exiting
            SaveTree demoTView, False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDemo = Nothing
End Sub

Private Sub mnuAdd2Root_click()
    AddRoot demoTView
End Sub
Private Sub AddRoot(tView As TreeView)
    '-----------------------------------
    '--add a node (no parent) to root --
    '-----------------------------------
    Dim NodX As Node
    Dim tCnt As Integer
    
    tCnt = IncrementCounter(tView)
    Set NodX = tView.Nodes.Add(, , "N" & tCnt, "Untitled(" & tCnt & ")", "iClosed", "iOpen")
    NodX.Checked = False
    NodX.Sorted = False
    TreeIsDirty = True
End Sub

Private Sub mnuAddChild_Click()
    'add a child node to the selected node
    AddChild demoTView
    
End Sub
Private Sub AddChild(tView As TreeView)
    '--------------------------------
    '-- Add a child node to parent --
    '--------------------------------
    If demoTView.SelectedItem Is Nothing Then Exit Sub
    
    Dim NodX As Node
    Dim tCnt As Integer, itemNdx As Integer
        
    itemNdx = tView.SelectedItem.Index
    tCnt = IncrementCounter(tView)
    Set NodX = tView.Nodes.Add(itemNdx, tvwChild, "N" & tCnt, "Untitled(" & tCnt & ")", "iClosed", "iOpen")
    NodX.Checked = False
    NodX.Sorted = False
    NodX.EnsureVisible
    TreeIsDirty = True

End Sub

Private Sub mnuRemoveNode_Click()
    '------------------------------
    '-- Remove the selected Node --
    '------------------------------
    If demoTView.SelectedItem Is Nothing Then Exit Sub
    
    demoTView.Nodes.Remove demoTView.SelectedItem.Index 'Removes the Node and any children it has
    cmdApplyChanges.Enabled = False
    ClearInfo

End Sub
Private Function GetIconNum(Str As String) As Integer
    '----------------------------------------
    '-- Returns the list index for an icon --
    '----------------------------------------
    GetIconNum = -1
    Dim x As Integer
    For x = 0 To cmbRegImg.ListCount - 1
        If Str = cmbRegImg.List(x) Then
            GetIconNum = x
        End If
    Next
End Function
Private Function GetParent(iNdx As Integer) As Integer
    '------------------------------------
    '-- Return the parent index if any --
    '------------------------------------
        
    If demoTView.Nodes(iNdx).Parent Is Nothing Then
        GetParent = 0
    Else
        GetParent = demoTView.Nodes(iNdx).Parent.Index
    End If

End Function
Private Function GetParentKey(iNdx As Integer) As String
    '----------------------------------------
    '-- Return the parent index key if any --
    '----------------------------------------
    If demoTView.Nodes(iNdx).Parent Is Nothing Then
        GetParentKey = "None"
    Else
        GetParentKey = demoTView.Nodes(iNdx).Parent.Key
    End If
    
End Function
Private Function IncrementCounter(tView As TreeView) As Integer
    '-------------------------------------------------------
    '-- Return the next index # for serializing node keys --
    '-------------------------------------------------------
    Dim cFile As String
    Dim ctr As CountInfo
    Dim F As Integer
    
    'create/read a counter file to prevent clashes in item keys
    cFile = getdir & tView.Name & ".dat" 'name is based on treeview.name to allow for multiple treeviews
    F = FreeFile
    Open cFile For Random As #F Len = Len(ctr) 'Random access will not generate an error if the file does not exist... it will just create the file
        If LOF(F) = 0 Then
            ctr.iCnt = 0
            Put #F, 1, ctr
        End If
        Get #F, 1, ctr
        IncrementCounter = ctr.iCnt + 1
        ctr.iCnt = IncrementCounter
        Put #F, 1, ctr
    Close #F
End Function


