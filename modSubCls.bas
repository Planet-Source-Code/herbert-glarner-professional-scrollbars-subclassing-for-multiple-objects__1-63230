Attribute VB_Name = "modSubCls"
Option Explicit

'FILE INFO
'---------
'Module:        modSubCls
'IPN prefix:    [e|t]msc
'Purpose:       Subclassing scrollbars for objects of gucScrWin controls

'Author:        Herbert Glarner
'Contact:       herbert.glarner@bluewin.ch
'Copyright:     (c) 2005 by Herbert Glarner
'               Freeware, provided you include credits and mail.
'               High risk of crashes if manipulated. READ COMMENTS! Use at own risk. No liability.



'PRIVATE CONSTANTS
'-----------------

'Window Procedure.
Private Const gclGWLWindProc As Long = -4&

'Intercepted Windows messages (Scrollbars messages).
Private Const gclWMHScroll As Long = &H114&
Private Const gclWMVScroll As Long = &H115&



'PRIVATE VARIABLES
'-----------------

'None, due to the multi-object usage. All variable data stored directly in
'the object windows.



'API DECLARATIONS
'----------------

'Installing/Removing a message hook to intercept scrollbar messages. Used to
'install an own Message Handler and to restore the original Windows Message
'Handler .
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Calling the original Windows message handler for messages of no particular
'interest (i.e. everything which we don't intercept).
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'Used to create illegal references instead of COM object references (to
'prevent controls from terminating due to an un-zero count)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'Storing/retrieving/removing variables stored for a specific windows handle.
'Details find in procedure "Hook".
Declare Function GetProp Lib "user32" Alias "GetPropA" _
    (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function SetProp Lib "user32" Alias "SetPropA" _
    (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
    (ByVal hWnd As Long, ByVal lpString As String) As Long



'ESTABLISHING AND REMOVING THE HOOK
'----------------------------------

'Establishing the Hook. Called by an object of the "gucScrWin" control.
Public Sub Hook(hWnd As Long, gswHooked As gucScrWin)
    Dim lObjectPtr As Long
    Dim lOrigWinProc As Long
    
Debug.Print "Hooking "; Hex$(hWnd)
    
    'We must avoid to hook an already hooked object. Therefore, we need a safe
    'place to store a pointer to the object (see below): We will need it for each
    'and every call to our object's "InterceptedWinMsg" method (called in this
    'modules "NewWinProc" procedure). We are a module, and used by several objects:
    'This makes global variables pretty meaningless (dangerous, actually). Thus
    'we use Window's properties storing mechanism, which lets us assign as many
    'variables to a window (by hWnd) as we want. After setting a variable with
    '"SetProp", we can retrieve it with "GetProp": no need for a complicated arrays
    'handling. - We call the variable holding aforementioned pointer "ObjPtr".
    
    'We check now, if the hook was already done.
    If GetProp(hWnd, "ObjPtr") = 0& Then
        'Remember the original Windows message handler: We pass every message not
        'of interest for our purpose back to it. We also need to reset the message
        'handler back to this address during deinstallation of the hook. If we
        'don't, at least the IDE will crash.
        lOrigWinProc = SetWindowLong(hWnd, gclGWLWindProc, AddressOf NewWinProc)
        
        'Store it as a property pertaining to this hWnd.
        SetProp hWnd, "OrigWinProc", lOrigWinProc

        'Instead of using 'Set goHooked = gswHooked', we store the pointer of
        'the object. The "Set" variant increases the object count, whereas just
        'obtaining an object pointer of course doesn't.
        lObjectPtr = ObjectToPtr(gswHooked)
        
        'Store it as well. (This is the variable tested in the start of this
        'procedure)
        SetProp hWnd, "ObjPtr", lObjectPtr
        
Debug.Print "ObjectToPtr "; Hex$(lObjectPtr)
Debug.Print "Hooked "; Hex$(hWnd)

    Else
Debug.Print "Already hooked: "; Hex$(hWnd)
    End If
End Sub

'Removing the established Hook.
Public Sub UnHook(hWnd As Long)
    Dim lObjectPtr As Long
    Dim lOrigWinProc As Long
    
Debug.Print "Unhooking "; Hex$(hWnd)
    
    'Get the stored object pointer.
    lObjectPtr = GetProp(hWnd, "ObjPtr")
    If lObjectPtr Then
        'Get the original windows message handler.
        lOrigWinProc = GetProp(hWnd, "OrigWinProc")
        
        'Restoring the original windows message handler.
        SetWindowLong hWnd, gclGWLWindProc, lOrigWinProc
        
        'We don't need the stored variables any longer.
        RemoveProp hWnd, "ObjPtr"
        RemoveProp hWnd, "OrigWinProc"
        
Debug.Print "Unhooked "; Hex$(hWnd)
    Else
Debug.Print "Was not hooked: "; Hex$(hWnd)
    End If
End Sub



'MESSAGE HANDLER
'---------------

'This is the new message handler, which will intercept the messages of interest
'for our problem. - Be careful her, and foremost: Be fast - windows waits for
'you to finish ;)
Private Function NewWinProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'The object for which we establish a hook. We need its reference to notify
    'it about intercepted Windows messages: this is done by calling the object's
    'public method "InterceptedWinMsg". - However, we do *not* store a validly
    'obtained COM reference into here, but an object pointer obtained illegally.
    'This variable is set just before we use it, by converting the invalid but
    'uncounted reference into a such.
    Dim oHooked As gucScrWin    'Goes out of scope at proc end: No count ;)
    Dim lObjectPtr As Long      'Was stored for the hWnd
    Dim lOrigWinProc As Long    'Uninteresting Windows messages
        
    'Intercept the messages of interest.
    If uMsg = gclWMHScroll Or uMsg = gclWMVScroll Then
        'Temporarily create an object from the stored object pointer.
        lObjectPtr = GetProp(hWnd, "ObjPtr")
        Set oHooked = PtrToObject(lObjectPtr)   'Goes out of scope at proc end.
    
        'Vertical scrollbar messages: Pass to object for handling.
        oHooked.InterceptedWinMsg hWnd, uMsg, wParam, lParam
    Else
        'Pass the other messages to windows, or they won't be processed at all!
        lOrigWinProc = GetProp(hWnd, "OrigWinProc")
        NewWinProc = CallWindowProc(lOrigWinProc, hWnd, uMsg, wParam, lParam)
    End If
End Function



'COM REFERENCE CIRCUMVENTION
'---------------------------

'Return the pointer for an object.
Private Property Get ObjectToPtr(ByRef gswHooked As gucScrWin) As Long
    ObjectToPtr = ObjPtr(gswHooked)
End Property

'Make a legal reference from an object pointer.
'DO NOT END WHILE WITHIN THIS PROCEDURE: THERE IS A 100% GUARANTEE FOR A CRASH.
Private Property Get PtrToObject(ByVal Ptr As Long) As gucScrWin
    Dim gswHooked As Object

    'Copy the pointer, thus creating an illegal, but uncounted interface.
    CopyMemory gswHooked, Ptr, 4&
    
    'We make the copy a legal reference now.
    Set PtrToObject = gswHooked
    
    'Clean up by getting rid of the illegal reference.
    CopyMemory gswHooked, 0&, 4&
End Property



