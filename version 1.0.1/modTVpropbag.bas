Attribute VB_Name = "modTVpropbag"
Option Explicit
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ TreeView Load/Save using propertybag demo                                             ++
'++ Ver. 1.0.1 Written by: CptnVic  29 Feb. 2008                                          ++
'++ Ver. 1.0.0 Written by: CptnVic  7 Feb. 2008                                           ++
'++ Copyright: Author retains ownership and his rights to use this code, however, you may ++
'++            use the code in your projects as you wish with no restriction provided     ++
'++            that you leave this copyright notice intact.                               ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Please read the attached readme.txt file                                              ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ The code in this module will save and restore the contents and most properties of a   ++
'++ treeview control.  At present, it saves and restores:                                 ++
'++   * All nodes (root and child nodes)            * Node Keys                           ++
'++   * Node icons (image and selected image)       * Node text                           ++
'++   * Node expanded state                         * Node tag                            ++
'++   * Node checked state                          * Node sorted state                   ++
'++ The subs will support saving/loading multiple treeviews                               ++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub LoadTree(tView As TreeView, Optional lFile As Variant)
    '----------------------------------------------------------------------------------
    '-- This sub loads the treeview entries from a file created by the SaveTree sub. --
    '----------------------------------------------------------------------------------
    Dim lF As Integer, numRecs As Integer, iCnt As Integer
    Dim lBag As New PropertyBag
    Dim lArray() As Byte
    
    'some variables for adding nodes to the treeview:
    Dim iNdx As Integer 'an integer representing the parent node (if any) treeview index
    Dim iKey As String 'a unique key for an entry
    Dim iTxt As String 'the text displayed in the treeview
    Dim iTag As String 'a good place to store just about anything... a filename/path for example
    Dim iIcon As String 'the key of an icon (in imagelist) to be displayed when item is NOT selected
    Dim iSIcon As String 'the key of an icon (in imagelist) to be displayed when item IS selected
    Dim isSorted As Boolean 'is node sorted?
    Dim iChecked As Boolean 'is node checked?
    Dim iExpanded As Boolean 'is node expanded?
    
    If IsMissing(lFile) Then
        'generate a unique file name based on the treeview name, this allows the sub to
        'process multiple treeviews - just passing the treeview's name.
        lFile = getdir & tView.Name & ".bag"
    End If
    lF = FreeFile
    Open lFile For Binary Access Read As #lF
        If LOF(lF) = 0 Then
            Close #lF
            Exit Sub
        End If
        tView.Parent.MousePointer = vbHourglass 'alert user to potential wait
        ReDim lArray(LOF(lF)) As Byte 'make array the size of the propertybag
        Get #lF, , lArray 'retrieve the property bag from file and place it in byte array
    Close #lF
    lBag.Contents = lArray 'load the byte array with the file (propertybag) contents
    tView.Nodes.Clear 'clear the treeview
    numRecs = lBag.ReadProperty("NumNodes") 'retrieve # of entries to process
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++ NOTE:  If you change the properties below (read with the lBag.ReadProperty),   ++
    '++        BE SURE to make similar changes in the SaveTree sub sBag.WriteProperty! ++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    For iCnt = 1 To numRecs
        'retrieve relevent info for this entry
        iNdx = lBag.ReadProperty("Parent" & iCnt) ' a 0 or index of parent
        iKey = lBag.ReadProperty("Key" & iCnt)
        iTxt = lBag.ReadProperty("Txt" & iCnt)
        iTag = lBag.ReadProperty("Tag" & iCnt)
        iIcon = lBag.ReadProperty("Icon" & iCnt)
        iSIcon = lBag.ReadProperty("IconSel" & iCnt)
        isSorted = lBag.ReadProperty("Sort" & iCnt)
        iChecked = lBag.ReadProperty("Check" & iCnt)
        iExpanded = lBag.ReadProperty("Expand" & iCnt)
        'add node to treeview based on iNdx value... it is either parentless or a child
        If iNdx = 0 Then
            LoadRootEntry tView, iKey, iTxt, iTag, iIcon, iSIcon, isSorted, iChecked, iExpanded
        Else
            LoadChildEntry tView, iNdx, iKey, iTxt, iTag, iIcon, iSIcon, isSorted, iChecked, iExpanded
        End If
    Next
    tView.Parent.MousePointer = vbDefault  'restore the mouse pointer
    'this stuff should be destroyed when the sub exits... but better safe than sorry?
    Erase lArray
    Set lBag = Nothing
End Sub
Public Sub SaveTree(tView As TreeView, Optional sFile As Variant, Optional DoBackUp As Boolean = False)
    '----------------------------------------------------------------------------------
    '-- This sub saves the treeview entries to a file named after the treeview...    --
    '-- or Optionally a filename you pass to the sub.                                --
    '----------------------------------------------------------------------------------
    Dim sF As Integer, iCnt As Integer
    Dim sBag As New PropertyBag
    Dim sArray() As Byte
    Dim NodX As Node
    
    tView.Parent.MousePointer = vbHourglass 'alert user to potential wait
    If IsMissing(sFile) Then
        sFile = getdir & tView.Name & ".bag"
    End If
    If DoBackUp Then
        BackUpTreeView CStr(sFile)
    End If
    'remove any existing file contents
    sF = FreeFile
    Open sFile For Output As #sF
    Close #sF
    
    sBag.WriteProperty "NumNodes", tView.Nodes.Count 'store # of items to be replaced on load
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++ NOTE:  If you change the properties below (saved with the sBag.WriteProperty), ++
    '++        BE SURE to make similar changes in the LoadTree sub lBag.WriteProperty! ++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        For iCnt = 1 To tView.Nodes.Count
            Set NodX = tView.Nodes(iCnt)
            'place relevent info for this entry into propertybag
            sBag.WriteProperty "Parent" & iCnt, sGetParent(tView, iCnt)  ' 0 or index of parent
            sBag.WriteProperty "Key" & iCnt, NodX.Key
            sBag.WriteProperty "Txt" & iCnt, NodX.Text
            sBag.WriteProperty "Tag" & iCnt, NodX.Tag
            sBag.WriteProperty "Icon" & iCnt, NodX.Image
            sBag.WriteProperty "IconSel" & iCnt, NodX.SelectedImage
            sBag.WriteProperty "Sort" & iCnt, NodX.Sorted
            sBag.WriteProperty "Check" & iCnt, NodX.Checked
            sBag.WriteProperty "Expand" & iCnt, NodX.Expanded
        Next
    'save the treeview info.
    sF = FreeFile
    Open sFile For Binary Access Write As #sF
        sArray = sBag.Contents 'place contents into byte array
    Put #sF, , sArray 'save byte array to file
    Close #sF
    tView.Parent.MousePointer = vbDefault 'restore the mouse pointer
    'this stuff should be destroyed when the sub exits... but better safe than sorry?
    Erase sArray
    Set sBag = Nothing
End Sub

Private Function sGetParent(Tgt As TreeView, iNdx As Integer) As Integer
    '------------------------------------------------------------------------------------
    '-- Return the parent index if any                                                 --
    '-- In Ver. 1.0 I used an error to set this property... a VERY bad coding practice.--
    '-- SADLY, I got the idea from MSDN, but ... this is a much better method.         --
    '------------------------------------------------------------------------------------
    If Tgt.Nodes(iNdx).Parent Is Nothing Then
        sGetParent = 0
    Else
        sGetParent = Tgt.Nodes(iNdx).Parent.Index
    End If
End Function

Private Sub LoadRootEntry(Obj As TreeView, id As String, Txt As String, iTag As String, rImg As String, sImg As String, iSorted As Boolean, itemChk As Boolean, eXpand As Boolean)
    '--------------------------------------------------------------------------
    '-- This sub adds a node at the root level.  This new node has no parent --
    '--------------------------------------------------------------------------
    Dim NodX As Node
    
    Set NodX = Obj.Nodes.Add(, , id, Txt, rImg, sImg) 'add the node to the treeview
    'set stuff that the add method doesn't do:
    NodX.Tag = iTag
    NodX.Sorted = iSorted
    NodX.Checked = itemChk
    NodX.Expanded = eXpand
    
End Sub
Private Sub LoadChildEntry(Obj As TreeView, OwnerNdx As Integer, idKey As String, Txt As String, iTag As String, iconReg As String, iconSel As String, iSorted As Boolean, itemChk As Boolean, eXpand As Boolean)
    '-------------------------------------------------------
    '-- This sub adds a child to the selected parent node --
    '-------------------------------------------------------
    Dim NodX As Node
    
    Set NodX = Obj.Nodes.Add(OwnerNdx, tvwChild, idKey, Txt, iconReg, iconSel) 'add the node to the treeview
    'set stuff that the add method doesn't do:
    NodX.Tag = iTag
    NodX.Sorted = iSorted
    NodX.Checked = itemChk
    NodX.Expanded = eXpand
    
End Sub
Private Sub BackUpTreeView(fileIn As String)
    Dim fileOut As String
    fileOut = getdir & Format(Now, "DDMMYYYY") & ".bag"
    FileCopy fileIn, fileOut 'copy the current file to the bak file
    '-- note: filecopy overwrites the existing bak file if you save more than once per day!
    '-- If that would be bad, change the file naming code for fileOut above.
    '-- I really only used this during de-bugging.
End Sub
Public Function getdir() As String
    '----------------------------------------------------
    '-- This function returns the application's path   --
    '-- and appends a "\" if necessary                 --
    '-- Change as desired if you want a different path --
    '----------------------------------------------------
    '-- p.s.  this is only public because I re-use it  --
    '-- by calling this function in the form module.   --
    '----------------------------------------------------
    getdir = App.Path
    If Right(getdir, 1) <> "\" Then getdir = getdir & "\"
End Function
