VERSION 5.00
Begin VB.UserControl gucScrWin 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   ScaleHeight     =   2415
   ScaleWidth      =   2880
End
Attribute VB_Name = "gucScrWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'FILE INFO
'---------
'Control:       gucScrWin
'IPN prefix:    [e|t]gsw
'Purpose:       User control with ative windows scrollbars, using modSubCls for subclassing

'Author:        Herbert Glarner
'Contact:       herbert.glarner@bluewin.ch
'Copyright:     (c) 2005 by Herbert Glarner
'               Freeware, provided you include credits and mail.
'               High risk of crashes if manipulated. READ COMMENTS! Use at own risk. No liability.



'USAGE
'-----
'(1) Use "ActivateScrollbars" to set the Hook (implemented here in Initialize).
'From now on, windows messages are intercepted and processed.
'
'(2) Assign the type of scrollbar(s) to be used (Horiz/Vert/Both) with "ActiveScrollbars".
'
'(3) Use "Min", "Max", "LargeChange" and "Value" to specify the scrollbars' properties.
'Until now, nothing was displayed: we defined just how the future scrollbar(s) will look like.
'These definitions are communicated to windows now.
'
'(4) Use "SetScrollbar" to inform Windows about what we want (once for each scrollbar).
'Time to display the scrollbar(s) now (both in one go, if both were defined).
'
'(5) Display the scrollbar(s) via "ShowScrollbars" (hide them via "HideScrollbars")
'The scrollbars are displayed and functional now, but VB does not receive the Windows
'Messages yet. To be able to do so, we set a Hook to intercept the messages of interest.
'
'(6) When done, use "DeactivateScrollbars" to release the Hook and restore the original handler.



'MULTIPLE USAGE
'--------------
'The Module stores no data globally pertaining to individual objects. Instead, the built-in
'storage facility is used to store data for the individual windows. (Check the module.)



'EXPOSED PROPERTIES
'------------------
'ActiveScrollbars       r/w     Horizontal, Vertical, or both scrollbars
'LargeChange            r/w     Property for one of the scrollbars
'Max                    r/w     Property for one of the scrollbars
'Min                    r/w     Property for one of the scrollbars
'Value                  r/w     Property for one of the scrollbars



'EXPOSED METHODS
'---------------
'HideScrollbars         Hide the scrollbars as per "ActiveScrollbars" settings
'SetScrollbar           Make Windows be aware about our scrollbar(s) and their settings
'ShowScrollbars         Display the scrollbars as per "ActiveScrollbars" settings
'InterceptedWinMsg      Called by the module for an intercepted Windows message



'ENUMERATIONS
'------------

'Type of scrollbar. Used in API calls.
Public Enum egswSBDefinition
    egswSBDHorizontal = 0&
    egswSBDVertical = 1&
    egswSBDBoth = 3&
End Enum

'Our properties allow setting the value for either oneof the scrollbars, but not
'for both together: we have an individual record of the "tgswScrollInfo" structure
'for each.
Public Enum egswSBOrientation
    egswSBOHorizontal = 0&
    egswSBOVertical = 1&
End Enum

'Intercepted Windows messages. Used in the 'public' prodedure "InterceptedWinMsg"
'which is called by the module when a Windows message of interest was intercepted.
Public Enum egswWindowsMessage
    'IDs for Scrollbars messages
    egswWMHScroll = &H114&
    egswWMVScroll = &H115&
End Enum

'The scrollbar notification types are delivered in the low word of the DWord
'"wParam" (that is an 'argument' of the called method "InterceptedWinMsg"). Use
'the private function "GetLoWord" to extract that word from wParam.
Public Enum egswSBNotification
    'Set scroll value to value - SmallChange
    egswSBNLineLeft = 0
    egswSBNLineUp = 0
    'Set scroll value to value + SmallChange
    egswSBNLineDown = 1
    egswSBNLineRight = 1
    'Set scroll value to value - LargeChange
    egswSBNPageLeft = 2
    egswSBNPageUp = 2
    'Set scroll value to value + LargeChange
    egswSBNPageRight = 3
    egswSBNPageDown = 3
    'Set scroll value to track position, Track Event if wanted
    egswSBNThumbTrack = 5       'while Tracking
    egswSBNThumbPosition = 4    'End of Tracking
    'Set scroll value to min
    egswSBNLeft = 6
    egswSBNTop = 6
    'Set scroll value to max
    egswSBNRight = 7
    egswSBNBottom = 7
    'Raise a Change Event
    egswSBNEndScroll = 8
End Enum

'Used in the "Mask" field of the structure "tgswScrollInfo".
Public Enum egswScrollInfoMask
    egswSIMRange = &H1
    egswSIMPage = &H2
    egswSIMPos = &H4
    egswSIMDisableNoScroll = &H8
    egswSIMTrackPos = &H10
    egswSIMAll = (egswSIMRange Or egswSIMPage Or egswSIMPos Or egswSIMTrackPos)
End Enum



'TYPES
'-----

'MS's SCROLLINFO structure. Used to set/retrieve scrollbar values.
Private Type tgswScrollInfo
    Size As Long                'Size of (this) structure
    Mask As egswScrollInfoMask  'Values to change
    Min As Long                 'Minimum value of the scrollbar
    Max As Long                 'Maximum valueof the scrollbar
    Page As Long                'What VB calls "LargeChange"
    Pos As Long                 'Current value
    TrackPos As Long            '[Is actually in HiWord of wParam]
End Type
'Note, that the actual maximal value of the scrollbar is actually equal to the
'structure's "Max" value plus its "Page" value.



'INTERNAL VARIABLES
'------------------

'Stores the active scrollbar(s). Use "ActiveScrollbars" to set/read this value.
Private glSBDefinition As egswSBDefinition

'We need a "tgswScrollInfo" record per scrollbar, i.e. one each for the
'horizontal (egswSBOHorizontal) and the vertical (egswSBOVertical) scrollbar.
Private rgswScrollInfo(egswSBOHorizontal To egswSBOVertical) As tgswScrollInfo

'Is True when a message hook was established via "ActivateScrollbars". If you
'terminate the program without removing the hook first (automatically done in
'the class' "Terminate" event, also possible via "DeactivateScrollbars").
Private gbHookSet As Boolean

'We're only raising a "Change" event if there is a new value. This variable holds
'the last value for which such an event was raised.
Private glLastEventValue As Long



'EXPOSED EVENTS
'--------------

'Raised when there is a new Value for the scrollbar.
Public Event Change(Scrollbar As egswSBOrientation, Value As Long)

'Add a "Scroll" event and in the method "ProcessScrollbar" do the appropriate
'changes if you need it separately. For my purposes, scrollbars always have
'the so-called "HotTracking" feature: what's moved is the new position, and
'that gets communicated by the "Change" event.



'API DECLARATIONS
'----------------

'Shows or hides a scrollbar
Private Declare Function ShowScrollBar Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal wBar As egswSBDefinition, _
    ByVal bShow As Boolean) As Long

'Sets the properties of a scrollbar
Private Declare Function SetScrollInfo Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal wBar As egswSBOrientation, _
    ByRef lpScrollInfo As tgswScrollInfo, ByVal bool As Boolean) As Long



'CONSTRUCTOR AND DESTRUCTOR
'--------------------------

Private Sub UserControl_Initialize()
Debug.Print "Initialize"
    Dim lSize As Long
    
    'The field "Size" of the two variables of structure type "tgswScrollInfo"
    'needs to be set once only: it won't change.
    lSize = Len(rgswScrollInfo(egswSBOHorizontal))  'Either one does the job
    rgswScrollInfo(egswSBOHorizontal).Size = lSize
    rgswScrollInfo(egswSBOVertical).Size = lSize
    
    rgswScrollInfo(egswSBOHorizontal).Mask = egswSIMAll
    rgswScrollInfo(egswSBOVertical).Mask = egswSIMAll
    
    'Installs the hook
    ActivateScrollbars
End Sub

Private Sub UserControl_Terminate()
Debug.Print "Terminate"
    'Removes the hook.
    DeactivateScrollbars
End Sub



'PUBLIC PROPERTIES
'-----------------

'Assign the type of scrollbar(s) to be displayed, read what type(s) were assigned.
Public Property Let ActiveScrollbars(BarsToDisplay As egswSBDefinition)
    glSBDefinition = BarsToDisplay
End Property
Public Property Get ActiveScrollbars() As egswSBDefinition
    ActiveScrollbars = glSBDefinition
End Property

'Assign/Read the scrollbar property Min/Max/Value/LargeChange for one of the two
'scrollbars (horizontal or vertical).
Public Property Let LargeChange(Scrollbar As egswSBOrientation, NewValue As Long)
    rgswScrollInfo(Scrollbar).Page = NewValue
End Property
Public Property Get LargeChange(Scrollbar As egswSBOrientation) As Long
    LargeChange = rgswScrollInfo(Scrollbar).Page
End Property

Public Property Let Max(Scrollbar As egswSBOrientation, NewMaximum As Long)
    rgswScrollInfo(Scrollbar).Max = NewMaximum
End Property
Public Property Get Max(Scrollbar As egswSBOrientation) As Long
    Max = rgswScrollInfo(Scrollbar).Max
End Property

Public Property Let Min(Scrollbar As egswSBOrientation, NewMinimum As Long)
    rgswScrollInfo(Scrollbar).Min = NewMinimum
End Property
Public Property Get Min(Scrollbar As egswSBOrientation) As Long
    Min = rgswScrollInfo(Scrollbar).Min
End Property

Public Property Let Value(Scrollbar As egswSBOrientation, NewValue As Long)
    rgswScrollInfo(Scrollbar).Pos = NewValue
End Property
Public Property Get Value(Scrollbar As egswSBOrientation) As Long
    Value = rgswScrollInfo(Scrollbar).Pos
End Property



'PUBLIC METHODS
'--------------

'Communicating the desired settings (Min, Max, Value, LargeChange) to Windows.
Public Sub SetScrollbar(Scrollbar As egswSBOrientation)
    SetScrollInfo hWnd, Scrollbar, rgswScrollInfo(Scrollbar), True
End Sub

'Showing the scrollbars as defined in "ActiveScrollbars".
Public Sub ShowScrollbars()
    ShowScrollBar hWnd, glSBDefinition, True
End Sub

'Hide the scrollbars as defined in "ActiveScrollbars".
Public Sub HideScrollbars()
    ShowScrollBar hWnd, glSBDefinition, False
End Sub



'CALLS FROM MODULE
'-----------------

'This method is called from the module, when a Windows message was intercepted
'which is of interested for our purpose. - Keep the code in here fast: Windows
'waits for its return before continuing to work.
Public Sub InterceptedWinMsg(ByVal hWnd As Long, ByVal uMsg As egswWindowsMessage, _
    ByVal wParam As Long, ByVal lParam As Long)

    'Process the interesting Windows messages, send the rest back to Windows.
    If uMsg = egswWMHScroll Then
        'Horizontal scrollbar messages
        ProcessScrollBar egswSBOHorizontal, GetLoWord(wParam), GetHiWord(wParam)
    ElseIf uMsg = egswWMVScroll Then
        'Vertical scrollbar messages
        ProcessScrollBar egswSBOVertical, GetLoWord(wParam), GetHiWord(wParam)
    Else
Debug.Print "Intercepting unhandled WM", uMsg
    End If
End Sub



'PRIVATE METHODS
'---------------

'Sets a Hook to intercept and process the Windows messages of interest. We MUST
'do this in a module, because the IDE is not able to process Callback functions
'for user controls.
Private Sub ActivateScrollbars()
    If Not gbHookSet Then
Debug.Print "Installing Hook"
        modSubCls.Hook hWnd, Me
        gbHookSet = True
    End If
End Sub

'Removes the established hook. this MUST be called before the program is left or
'the IDE, possibly the whole system will crash. Also called by the class'
'"Terminate" event.
Private Sub DeactivateScrollbars()
    If gbHookSet Then
Debug.Print "Removing Hook"
        modSubCls.UnHook hWnd
        gbHookSet = False
    End If
End Sub

'Processing a scrollbar notification. Called by InterceptedWinMsg for either of
'the two scrollbar orientations (Scrollbar tells for which).
Private Sub ProcessScrollBar(Scrollbar As egswSBOrientation, _
    Notification As egswSBNotification, nPos As Long)
    
    Dim lValue As Long
    Dim eMask As egswScrollInfoMask
    Dim lEffMax As Long
    
    With rgswScrollInfo(Scrollbar)
        'The other notifications all change the position (the 'value').
        Select Case Notification
            Case egswSBNThumbTrack, egswSBNThumbPosition
                'Set scroll value to track position. Here, the scroll position
                'is provided in nPos (ex the Hi Word of wParam).
                lValue = nPos
            Case egswSBNLineUp      'also egswSBNLineLeft
                'Set scroll value to value - SmallChange
                lValue = .Pos - 1&
                If lValue < .Min Then lValue = .Min
            Case egswSBNLineDown    'also egswSBNLineRight
                'Set scroll value to value + SmallChange
                lValue = .Pos + 1&
                lEffMax = .Max - .Page + 1&
                If lValue > lEffMax Then lValue = lEffMax
            Case egswSBNPageUp      'also egswSBNPageLeft
                'Set scroll value to value - LargeChange
                lValue = .Pos - .Page
                If lValue < .Min Then lValue = .Min
            Case egswSBNPageDown    'also egswSBNPageRight
                'Set scroll value to value + LargeChange
                lValue = .Pos + .Page
                lEffMax = .Max - .Page + 1&
                If lValue > lEffMax Then lValue = lEffMax
            Case egswSBNTop         'also egswSBNLeft
                'Set scroll value to min
                lValue = .Min
            Case egswSBNBottom      'also egswSBNRight
                'Set scroll value to max
                lValue = .Max
        End Select
        
        'Provide the new values for Windows (not for egswSBNEndScroll)
        If Notification <> egswSBNEndScroll Then
            .Pos = lValue
            rgswScrollInfo(Scrollbar).Mask = egswSIMAll
            SetScrollbar Scrollbar
        End If
        
        '"glLastEventValue" holds the last value for which a "Change" event was
        'raised. A new event is raised only when it differs from the last event.
        'If you don't want hot tracking, use "egswSBNEndScroll" to raise a "Change"
        'event and "egswSBNThumbTrack" to raise a "Scroll" event.
        If glLastEventValue <> .Pos Then
            RaiseEvent Change(Scrollbar, .Pos)
            glLastEventValue = .Pos
        End If
    End With
End Sub

'Extracting the High Word of a DWord.
Private Function GetHiWord(ByVal DWord As Long) As Long
    GetHiWord = (DWord And &HFFFF0000) \ &H10000
End Function

'Extracting the Low Word of a DWord.
Private Function GetLoWord(ByVal DWord As Long) As Long
    DWord = DWord And &HFFFF&
    If DWord > 32767 Then GetLoWord = DWord - 65536 Else GetLoWord = DWord
End Function

