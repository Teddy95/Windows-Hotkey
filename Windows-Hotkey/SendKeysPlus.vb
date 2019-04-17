Imports System.ComponentModel
Imports System.Configuration
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Threading
Imports System.Windows.Forms


#If 1 = 0 Then
									*************************
									********* NOTES *********
									*************************

	In order to compile correctly, make sure you have these references added to your project:

						System
						System.configuration
						System.Core
						System.Data
						System.Data.DataSetExtensions
						System.Deployment
						System.Drawing
						System.Windows.Forms

	And these namespaces:

						Microsoft.VisualBasic
						System
						System.Collections
						System.Collections.Generic
						System.Data
						System.Diagnostics

	If you find that a key is missing, add it to the code (look for the tag _CODES_) with a instruction
	like:
							AddScanCode(Keys.xxxx, make_code, break_code)

	Have a look to that section right now and make sure all the keys you need are present. These are the instructions to add any other you need:
	
	You need to find the "make" and "break" codes. Last one is easy:
					break_code = make_code + &h80   (usually)
	So the only unknown parameter is "make_code". I found this page useful to locate it:

					http://www.quadibloc.com/comp/scan.htm

			It shows a keyboard layout with scan codes ("make_codes") that seems to work with occident keyboards:

			 ---     ---------------   ---------------   ---------------   -----------
			| 01|   | 3B| 3C| 3D| 3E| | 3F| 40| 41| 42| | 43| 44| 57| 58| |+37|+46|+45| 
			 ---     ---------------   ---------------   ---------------   -----------

			 -----------------------------------------------------------   -----------   ---------------
			| 29| 02| 03| 04| 05| 06| 07| 08| 09| 0A| 0B| 0C| 0D|     0E| |*52|*47|*49| |+45|+35|+37| 4A|
			|-----------------------------------------------------------| |-----------| |---------------|
			|   0F| 10| 11| 12| 13| 14| 15| 16| 17| 18| 19| 1A| 1B|   2B| |*53|*4F|*51| | 47| 48| 49|   |
			|-----------------------------------------------------------|  -----------  |-----------| 4E|
			|    3A| 1E| 1F| 20| 21| 22| 23| 24| 25| 26| 27| 28|      1C|               | 4B| 4C| 4D|   |
			|-----------------------------------------------------------|      ---      |---------------|
			|      2A| 2C| 2D| 2E| 2F| 30| 31| 32| 33| 34| 35|        36|     |*4C|     | 4F| 50| 51|   |
			|-----------------------------------------------------------|  -----------  |-----------|-1C|
			|   1D|-5B|   38|                       39|-38|-5C|-5D|  -1D| |*4B|*50|*4D| |     52| 53|   |
			 -----------------------------------------------------------   -----------   ---------------

	Code adapted from https://www.codeproject.com/Tips/310817/SendKeys-using-ScanCodes-to-Citrix
	Very grateful to the author!

#End If



Public Delegate Function CallBack(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

Namespace Utils


    ''' <summary>Provides methods for sending keystrokes to an application.</summary>
    ''' <filterpriority>2</filterpriority>
    Public Class SendKeysPlus

        Private Structure MOUSEINPUT
            Public dx As Integer
            Public dy As Integer
            Public mouseData As Integer
            Public dwFlags As Integer
            Public time As Integer
            Public dwExtraInfo As IntPtr
        End Structure

        Private Structure KEYBDINPUT
            Public wVk As Short
            Public wScan As Short
            Public dwFlags As Integer
            Public time As Integer
            Public dwExtraInfo As IntPtr
        End Structure

        Private Structure HARDWAREINPUT
            Public uMsg As Integer
            Public wParamL As Short
            Public wParamH As Short
        End Structure

        <StructLayout(LayoutKind.Explicit)>
        Private Structure INPUT
            <FieldOffset(0)>
            Public type As Integer
            <FieldOffset(4)>
            Public mi As MOUSEINPUT
            <FieldOffset(4)>
            Public ki As KEYBDINPUT
            <FieldOffset(4)>
            Public hi As HARDWAREINPUT
        End Structure

        Structure EVENTMSG
            Public message As UInt32
            Public paramL As UInt32
            Public paramH As UInt32
            Public time As UInt32
            Public hwnd As IntPtr
        End Structure

        Private Declare Function SendInput Lib "user32" (ByVal nInputs As Integer, ByVal pInputs() As INPUT, ByVal cbSize As Integer) As Integer

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function SetKeyboardState(ByVal lpKeyState() As Byte) As Boolean
        End Function

        <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
        Public Overloads Shared Function SetWindowsHookEx _
          (ByVal idHook As Integer, ByVal HookProc As CallBack,
           ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Public Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Shared Function GetAsyncKeyState(ByVal vKey As System.Windows.Forms.Keys) As Short
        End Function

        <DllImport("kernel32.dll")>
        Private Shared Function GetTickCount() As UInteger
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function VkKeyScan(ByVal ch As Char) As Short
        End Function


        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function OemKeyScan(ByVal ch As Char) As Integer
        End Function

        Private Declare Function GetKeyboardState Lib "user32" (ByVal keyState() As Byte) As Boolean
        Private Declare Function BlockInput Lib "user32" (ByVal fBlockIt As Boolean) As Boolean

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        End Function

        Private Class ScanCodeValue
            Private _Key As Byte
            Private _MakeCode As Byte
            Private _BreakCode As Byte

            Public Sub New(ByVal Key As Keys, ByVal MakeCode As Integer, ByVal Breakcode As Integer)
                _Key = Key
                _MakeCode = MakeCode
                _BreakCode = Breakcode
            End Sub

            Public ReadOnly Property Key() As Keys
                Get
                    Return _Key
                End Get
            End Property

            Public ReadOnly Property MakeCode() As Byte
                Get
                    Return _MakeCode
                End Get
            End Property

            Public ReadOnly Property BreakCode() As Byte
                Get
                    Return _BreakCode
                End Get
            End Property
        End Class

        Private Shared _ScanCodeTable As SortedList(Of Keys, ScanCodeValue)
        Private Shared ReadOnly Property ScanCodeTable() As SortedList(Of Keys, ScanCodeValue)
            Get
                If _ScanCodeTable Is Nothing Then
                    '_ScanCodeTable.Add(Keys., New ScanCodeValue(Keys., &H, &H))
                    _ScanCodeTable = New SortedList(Of Keys, ScanCodeValue)
                    ' _CODES_
                    AddScanCode(Keys.Menu, &H38, &HB8)
                    AddScanCode(Keys.ControlKey, &H1D, &H9D)
                    AddScanCode(Keys.ShiftKey, &H2A, &HAA)
                    AddScanCode(Keys.Back, &HE, &H8E)
                    AddScanCode(Keys.CapsLock, &H3A, &HBA)
                    AddScanCode(Keys.Enter, &H1C, &H9C)
                    AddScanCode(Keys.Escape, &H1, &H81)
                    AddScanCode(Keys.LMenu, &H38, &HB8)
                    AddScanCode(Keys.LControlKey, &H1D, &H9D)
                    AddScanCode(Keys.LShiftKey, &H2A, &HAA)
                    AddScanCode(Keys.NumLock, &H45, &HC5)
                    AddScanCode(Keys.RShiftKey, &H36, &HB6)
                    AddScanCode(Keys.Scroll, &H46, &HC6)
                    AddScanCode(Keys.Space, &H39, &HB9)
                    AddScanCode(Keys.Attn, &H54, &HD4)
                    AddScanCode(Keys.Tab, &HF, &H8F)
                    AddScanCode(Keys.A, &H1E, &H9E)
                    AddScanCode(Keys.B, &H30, &HB0)
                    AddScanCode(Keys.C, &H2E, &HAE)
                    AddScanCode(Keys.D, &H20, &HA0)
                    AddScanCode(Keys.E, &H12, &H92)
                    AddScanCode(Keys.F, &H21, &HA1)
                    AddScanCode(Keys.G, &H22, &HA2)
                    AddScanCode(Keys.H, &H23, &HA3)
                    AddScanCode(Keys.I, &H17, &H97)
                    AddScanCode(Keys.J, &H24, &HA4)
                    AddScanCode(Keys.K, &H25, &HA5)
                    AddScanCode(Keys.L, &H26, &HA6)
                    AddScanCode(Keys.M, &H32, &HB2)
                    AddScanCode(Keys.N, &H31, &HB1)
                    AddScanCode(Keys.O, &H18, &H98)
                    AddScanCode(Keys.P, &H19, &H99)
                    AddScanCode(Keys.Q, &H10, &H90)
                    AddScanCode(Keys.R, &H13, &H93)
                    AddScanCode(Keys.S, &H1F, &H9F)
                    AddScanCode(Keys.T, &H14, &H94)
                    AddScanCode(Keys.U, &H16, &H96)
                    AddScanCode(Keys.V, &H2F, &HAF)
                    AddScanCode(Keys.W, &H11, &H91)
                    AddScanCode(Keys.X, &H2D, &HAD)
                    AddScanCode(Keys.Y, &H15, &H95)
                    AddScanCode(Keys.Z, &H2C, &HAC)
                    AddScanCode(Keys.F1, &H3B, &HBB)
                    AddScanCode(Keys.F2, &H3C, &HBC)
                    AddScanCode(Keys.F3, &H3D, &HBD)
                    AddScanCode(Keys.F4, &H3E, &HBE)
                    AddScanCode(Keys.F7, &H41, &HC1)
                    AddScanCode(Keys.F5, &H3F, &HBF)
                    AddScanCode(Keys.F6, &H40, &HC0)
                    AddScanCode(Keys.F8, &H42, &HC2)
                    AddScanCode(Keys.F9, &H43, &HC3)
                    AddScanCode(Keys.F10, &H44, &HC4)
                    AddScanCode(Keys.F11, &H57, &HD7)
                    AddScanCode(Keys.F12, &H58, &HD8)
                    AddScanCode(Keys.D1, &H2, &H82)
                    AddScanCode(Keys.D2, &H3, &H83)
                    AddScanCode(Keys.D3, &H4, &H84)
                    AddScanCode(Keys.D4, &H5, &H85)
                    AddScanCode(Keys.D5, &H6, &H86)
                    AddScanCode(Keys.D6, &H7, &H87)
                    AddScanCode(Keys.D7, &H8, &H88)
                    AddScanCode(Keys.D8, &H9, &H89)
                    AddScanCode(Keys.D9, &HA, &H8A)
                    AddScanCode(Keys.D0, &HB, &H8B)
                    AddScanCode(Keys.OemMinus, &HC, &H8C)
                    AddScanCode(Keys.Oemplus, &HD, &H8D)
                    AddScanCode(Keys.OemOpenBrackets, &H1A, &H9A)
                    AddScanCode(Keys.OemCloseBrackets, &H1B, &H9B)
                    AddScanCode(Keys.OemSemicolon, &H27, &HA7)
                    AddScanCode(Keys.OemQuotes, &H29, &HA9)
                    AddScanCode(Keys.OemPipe, &H2B, &HAB)
                    AddScanCode(Keys.Oemcomma, &H33, &HB3)
                    AddScanCode(Keys.OemPeriod, &H34, &HB4)
                    AddScanCode(Keys.OemQuestion, &H35, &HB5)
                    AddScanCode(Keys.NumPad0, &H52, &HD2)
                    AddScanCode(Keys.NumPad1, &H4F, &HCF)
                    AddScanCode(Keys.NumPad2, &H50, &HD0)
                    AddScanCode(Keys.NumPad3, &H51, &HD1)
                    AddScanCode(Keys.NumPad4, &H4B, &HCB)
                    AddScanCode(Keys.NumPad5, &H4C, &HCC)
                    AddScanCode(Keys.NumPad6, &H4D, &HCD)
                    AddScanCode(Keys.NumPad7, &H47, &HC7)
                    AddScanCode(Keys.NumPad8, &H48, &HC8)
                    AddScanCode(Keys.NumPad9, &H49, &HC9)
                    AddScanCode(Keys.Decimal, &H53, &HD3)
                    AddScanCode(Keys.Multiply, &H37, &HB7)
                    AddScanCode(Keys.Subtract, &H4A, &HCA)
                    AddScanCode(Keys.Add, &H4E, &HCE)
                    AddScanCode(Keys.PageUp, &H49, &HC9)
                    AddScanCode(Keys.PageDown, &H51, &HD1)
                    AddScanCode(Keys.Down, &H50, &HD0)
                    AddScanCode(Keys.Up, &H4C, &HCC)
                End If
                Return _ScanCodeTable
            End Get
        End Property


        Private Shared Sub AddScanCode(ByVal Key As Keys, ByVal Make As Byte, ByVal Break As Byte)
            _ScanCodeTable.Add(Key, New ScanCodeValue(Key, Make, Break))
        End Sub


        ' Methods
        Shared Sub New()
            AddHandler Application.ThreadExit, New EventHandler(AddressOf SendKeysPlus.OnThreadExit)
            SendKeysPlus.messageWindow = New SKWindow
            SendKeysPlus.messageWindow.CreateControl()
        End Sub

        Private Sub New()
        End Sub

        Private Shared Sub AddCancelModifiersForPreviousEvents(ByVal previousEvents As Queue)
            If (Not previousEvents Is Nothing) Then
                Dim flag As Boolean = False
                Dim flag2 As Boolean = False
                Dim flag3 As Boolean = False
                'This Assumes that all of the previousEvents
                'have the same ScanCodeMode which they should
                Dim myScanCodeMode As ScanCodeModes = ScanCodeModes.None
                Do While (previousEvents.Count > 0)
                    Dim flag4 As Boolean
                    Dim event2 As SKEvent = DirectCast(previousEvents.Dequeue, SKEvent)
                    myScanCodeMode = event2.ScanCodeMode
                    If ((event2.wm = &H101) OrElse (event2.wm = &H105)) Then
                        flag4 = False
                    Else
                        If ((event2.wm <> &H100) AndAlso (event2.wm <> 260)) Then
                            Continue Do
                        End If
                        flag4 = True
                    End If
                    If (event2.paramL = &H10) Then
                        flag = flag4
                    Else
                        If (event2.paramL = &H11) Then
                            flag2 = flag4
                            Continue Do
                        End If
                        If (event2.paramL = &H12) Then
                            flag3 = flag4
                        End If
                    End If
                Loop
                If flag Then
                    SendKeysPlus.AddEvent(New SKEvent(&H101, &H10,
                     False, IntPtr.Zero, myScanCodeMode))  'Shift Key

                ElseIf flag2 Then
                    SendKeysPlus.AddEvent(New SKEvent(&H101, &H11,
                     False, IntPtr.Zero, myScanCodeMode)) 'Ctrl Key

                ElseIf flag3 Then
                    SendKeysPlus.AddEvent(New SKEvent(&H105, &H12, False,
                      IntPtr.Zero, myScanCodeMode)) ' Alt Key
                End If
            End If
        End Sub

        Private Shared Function GetScanCode(ByVal key As Keys,
          ByVal ScanCodeMode As ScanCodeModes) As Integer
            Dim myScanCode As Integer = 0
            Select Case ScanCodeMode
                Case ScanCodeModes.None
                    myScanCode = 0
                Case ScanCodeModes.Break
                    Dim myScanCodeValue As ScanCodeValue = Nothing
                    If ScanCodeTable.TryGetValue(key, myScanCodeValue) Then
                        If myScanCodeValue IsNot Nothing Then
                            myScanCode = myScanCodeValue.BreakCode
                        Else
                            myScanCode = 0
                        End If
                    End If
                Case ScanCodeModes.Make
                    Dim myScanCodeValue As ScanCodeValue = Nothing
                    If ScanCodeTable.TryGetValue(key, myScanCodeValue) Then
                        If myScanCodeValue IsNot Nothing Then
                            myScanCode = myScanCodeValue.MakeCode
                        Else
                            myScanCode = 0
                        End If
                    End If
            End Select
            Return myScanCode
        End Function

        Private Shared Sub AddEvent(ByVal skevent As SKEvent)
            If (SendKeysPlus.events Is Nothing) Then
                SendKeysPlus.events = New Queue
            End If
            SendKeysPlus.events.Enqueue(skevent)
        End Sub

        Private Shared Sub AddMsgsForVK(ByVal vk As Integer, ByVal repeat _
          As Integer, ByVal altnoctrldown As Boolean,
          ByVal hwnd As IntPtr, ByVal ScanCodeMode As ScanCodeModes)
            Dim i As Integer
            For i = 0 To repeat - 1
                SendKeysPlus.AddEvent(New SKEvent(If(altnoctrldown,
                  &H104, &H100), vk, SendKeysPlus.fStartNewChar, hwnd, ScanCodeMode))
                SendKeysPlus.AddEvent(New SKEvent(If(altnoctrldown,
                  &H105, &H101), vk, SendKeysPlus.fStartNewChar, hwnd, ScanCodeMode))
            Next i
        End Sub

        Private Shared Function AddSimpleKey(ByVal character As Char,
          ByVal repeat As Integer, ByVal hwnd As IntPtr, ByVal haveKeys _
          As Integer(), ByVal fStartNewChar As Boolean, ByVal cGrp As _
          Integer, ByVal ScanCodeMode As ScanCodeModes) As Boolean

            Dim num As Integer = VkKeyScan(character)
            If (num <> -1) Then
                If ((haveKeys(0) = 0) AndAlso ((num And &H100) <> 0)) Then
                    SendKeysPlus.AddEvent(New SKEvent(&H100, &H10, fStartNewChar, hwnd, ScanCodeMode))
                    fStartNewChar = False
                    haveKeys(0) = 10
                End If
                If ((haveKeys(1) = 0) AndAlso ((num And &H200) <> 0)) Then
                    SendKeysPlus.AddEvent(New SKEvent(&H100, &H11, fStartNewChar, hwnd, ScanCodeMode))
                    fStartNewChar = False
                    haveKeys(1) = 10
                End If
                If ((haveKeys(2) = 0) AndAlso ((num And &H400) <> 0)) Then
                    SendKeysPlus.AddEvent(New SKEvent(&H100, &H12, fStartNewChar, hwnd, ScanCodeMode))
                    fStartNewChar = False
                    haveKeys(2) = 10
                End If
                SendKeysPlus.AddMsgsForVK((num And &HFF), repeat, ((haveKeys(2) > 0) _
                 AndAlso (haveKeys(1) = 0)), hwnd, ScanCodeMode)
                SendKeysPlus.CancelMods(haveKeys, 10, hwnd, ScanCodeMode)
            Else
                'Dim num2 As Integer = _
                '  SafeNativeMethods.OemKeyScan(CShort((CByte(ChrW(255)) And CByte(character))))
                Dim num2 As Integer = OemKeyScan(ChrW((&HFF00 And Asc(character))))
                Dim i As Integer
                For i = 0 To repeat - 1
                    SendKeysPlus.AddEvent(New SKEvent(&H102, AscW(character),
                       (num2 And &HFFFF), hwnd, ScanCodeMode))
                Next i
            End If
            If (cGrp <> 0) Then
                fStartNewChar = True
            End If
            Return fStartNewChar
        End Function

        Private Shared Sub CancelMods(ByVal haveKeys As Integer(),
          ByVal level As Integer, ByVal hwnd As IntPtr, ByVal ScanCodeMode As ScanCodeModes)
            If (haveKeys(0) = level) Then
                SendKeysPlus.AddEvent(New SKEvent(&H101, &H10, False, hwnd, ScanCodeMode))
                haveKeys(0) = 0
            End If
            If (haveKeys(1) = level) Then
                SendKeysPlus.AddEvent(New SKEvent(&H101, &H11, False, hwnd, ScanCodeMode))
                haveKeys(1) = 0
            End If
            If (haveKeys(2) = level) Then
                SendKeysPlus.AddEvent(New SKEvent(&H105, &H12, False, hwnd, ScanCodeMode))
                haveKeys(2) = 0
            End If
        End Sub

        Private Shared Sub CheckGlobalKeys(ByVal skEvent As SKEvent)
            If (skEvent.wm = &H100) Then
                Select Case skEvent.paramL
                    Case 20
                        SendKeysPlus.capslockChanged = Not SendKeysPlus.capslockChanged
                        Return
                    Case &H15
                        SendKeysPlus.kanaChanged = Not SendKeysPlus.kanaChanged
                        Return
                    Case &H90
                        SendKeysPlus.numlockChanged = Not SendKeysPlus.numlockChanged
                        Return
                    Case &H91
                        SendKeysPlus.scrollLockChanged = Not SendKeysPlus.scrollLockChanged
                        Return
                    Case Else
                        Return
                End Select
            End If
        End Sub

        Private Shared Sub ClearGlobalKeys()
            SendKeysPlus.capslockChanged = False
            SendKeysPlus.numlockChanged = False
            SendKeysPlus.scrollLockChanged = False
            SendKeysPlus.kanaChanged = False
        End Sub

        Private Shared Sub ClearKeyboardState()
            Dim keyboardState As Byte() = SendKeysPlus.MyGetKeyboardState
            keyboardState(20) = 0
            keyboardState(&H90) = 0
            keyboardState(&H91) = 0
            SendKeysPlus.MySetKeyboardState(keyboardState)
        End Sub

        Private Shared Function EmptyHookCallback(ByVal code As Integer,
          ByVal wparam As IntPtr, ByVal lparam As IntPtr) As IntPtr
            Return IntPtr.Zero
        End Function

        ''' <summary>Processes all the Windows messages currently in the message queue.</summary>
        ''' <filterpriority>1</filterpriority>
        ''' <PermissionSet><IPermission
        '''    class="System.Security.Permissions.EnvironmentPermission, 
        '''    mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Unrestricted="true" />
        ''' <IPermission class="System.Security.Permissions.FileIOPermission, 
        '''    mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Unrestricted="true" />
        '''  <IPermission class="System.Security.Permissions.SecurityPermission, 
        '''     mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''     version="1" Flags="UnmanagedCode, ControlEvidence" />
        ''' <IPermission class="System.Security.Permissions.UIPermission, 
        '''    mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Unrestricted="true" />
        ''' <IPermission class="System.Diagnostics.PerformanceCounterPermission, 
        '''     System, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        ''' version="1" Unrestricted="true" /></PermissionSet>
        Public Shared Sub Flush()
            Application.DoEvents()
            Do While ((Not SendKeysPlus.events Is Nothing) AndAlso (SendKeysPlus.events.Count > 0))
                Application.DoEvents()
            Loop
        End Sub

        Private Shared Function MyGetKeyboardState() As Byte()
            Dim keystate As Byte() = New Byte(&H100 - 1) {}
            GetKeyboardState(keystate)
            Return keystate
        End Function

        Private Shared Sub InstallHook()
            If (SendKeysPlus.hhook = IntPtr.Zero) Then
                Dim proc1 As New SendKeysHookProc
                SendKeysPlus.hook = New CallBack(AddressOf proc1.Callback)
                SendKeysPlus.stopHook = False
                SendKeysPlus.hhook = SetWindowsHookEx(1, SendKeysPlus.hook,
                New HandleRef(Nothing, GetModuleHandle(Nothing)), 0)
                If (SendKeysPlus.hhook = IntPtr.Zero) Then
                    Throw New SecurityException("SendKeysHookFailed")
                End If
            End If
        End Sub

        Private Shared Function IsExtendedKey(ByVal skEvent As SKEvent) As Boolean
            If (((((skEvent.paramL <> &H26) AndAlso (skEvent.paramL <> 40)) _
             AndAlso ((skEvent.paramL <> &H25) AndAlso (skEvent.paramL <> &H27))) _
             AndAlso (((skEvent.paramL <> &H21) AndAlso (skEvent.paramL <> &H22)) _
             AndAlso ((skEvent.paramL <> &H24) AndAlso
             (skEvent.paramL <> &H23)))) AndAlso (skEvent.paramL <> &H2D)) Then
                Return (skEvent.paramL = &H2E)
            End If
            Return True
        End Function

        Private Shared Sub JournalCancel()
            If (SendKeysPlus.hhook <> IntPtr.Zero) Then
                SendKeysPlus.stopHook = False
                If (Not SendKeysPlus.events Is Nothing) Then
                    SendKeysPlus.events.Clear()
                End If
                SendKeysPlus.hhook = IntPtr.Zero
            End If
        End Sub

        Private Shared Sub LoadSendMethodFromConfig()
            If Not SendKeysPlus.sendMethod.HasValue Then
                SendKeysPlus.sendMethod = 1
                Try
                    Dim str As String = ConfigurationManager.AppSettings.Get("SendKeys")
                    If str.Equals("JournalHook", StringComparison.OrdinalIgnoreCase) Then
                        SendKeysPlus.sendMethod = 2
                    ElseIf str.Equals("SendInput", StringComparison.OrdinalIgnoreCase) Then
                        SendKeysPlus.sendMethod = 3
                    End If
                Catch
                End Try
            End If
        End Sub

        Public Shared Sub SetSendMethod(ByVal SendMethod As SendMethodTypes)
            SendKeysPlus.sendMethod = SendMethod
        End Sub

        Private Shared Function MatchKeyword(ByVal keyword As String) As Integer
            Dim i As Integer
            For i = 0 To SendKeysPlus.keywords.Length - 1
                If String.Equals(SendKeysPlus.keywords(i).keyword,
                 keyword, StringComparison.OrdinalIgnoreCase) Then
                    Return SendKeysPlus.keywords(i).vk
                End If
            Next i
            Return -1
        End Function

        Private Shared Sub OnThreadExit(ByVal sender As Object, ByVal e As EventArgs)
            Try
                SendKeysPlus.UninstallJournalingHook()
            Catch
            End Try
        End Sub

        Private Shared Sub ParseKeys(ByVal keys As String,
          ByVal hwnd As IntPtr, ByVal ScanCodeMode As ScanCodeModes)
            Dim num As Integer = 0
            Dim haveKeys As Integer() = New Integer(3 - 1) {}
            Dim cGrp As Integer = 0
            SendKeysPlus.fStartNewChar = True
            Dim length As Integer = keys.Length
            Do While (num < length)
                Dim num6 As Integer
                Dim num7 As Integer
                Dim repeat As Integer = 1
                Dim ch As Char = keys.Chars(num)
                Dim vk As Integer = 0
                Select Case ch
                    Case "%"c
                        If (haveKeys(2) <> 0) Then
                            Throw New ArgumentException("InvalidSendKeysString")
                        End If
                        GoTo Label_03C9
                    Case "("c
                        cGrp += 1
                        If (cGrp > 3) Then
                            Throw New ArgumentException("SendKeysNestingError")
                        End If
                        GoTo Label_0414
                    Case ")"c
                        If (cGrp < 1) Then
                            Throw New ArgumentException("InvalidSendKeysString")
                        End If
                        GoTo Label_045A
                    Case "+"c
                        If (haveKeys(0) <> 0) Then
                            Throw New ArgumentException("InvalidSendKeysString")
                        End If
                        GoTo Label_0333
                    Case "^"c
                        If (haveKeys(1) <> 0) Then
                            Throw New ArgumentException("InvalidSendKeysString")
                        End If
                        SendKeysPlus.AddEvent(New SKEvent(&H100, &H11,
                           SendKeysPlus.fStartNewChar, hwnd, ScanCodeMode))
                        SendKeysPlus.fStartNewChar = False
                        haveKeys(1) = 10
                        GoTo Label_04AB
                    Case "{"c
                        num6 = (num + 1)
                        If (((num6 + 1) >= length) OrElse (keys.Chars(num6) <> "}"c)) Then
                            GoTo Label_00EB
                        End If
                        num7 = (num6 + 1)
                        GoTo Label_00C7
                    Case "}"c
                        Throw New ArgumentException("InvalidSendKeysString")
                    Case "~"c
                        vk = 13
                        SendKeysPlus.AddMsgsForVK(vk, repeat,
                         ((haveKeys(2) > 0) AndAlso (haveKeys(1) = 0)), hwnd, ScanCodeMode)
                        GoTo Label_04AB
                    Case Else
                        SendKeysPlus.fStartNewChar =
                          SendKeysPlus.AddSimpleKey(keys.Chars(num), repeat, hwnd,
                          haveKeys, SendKeysPlus.fStartNewChar, cGrp, ScanCodeMode)
                        GoTo Label_04AB
                End Select
Label_00C1:
                num7 += 1
Label_00C7:
                If ((num7 < length) AndAlso (keys.Chars(num7) <> "}"c)) Then
                    GoTo Label_00C1
                End If
                If (num7 < length) Then
                    num6 += 1
                End If
Label_00EB:
                Do While (((num6 < length) AndAlso (keys.Chars(num6) <> "}"c)) _
                 AndAlso Not Char.IsWhiteSpace(keys.Chars(num6)))
                    num6 += 1
                Loop
                If (num6 >= length) Then
                    Throw New ArgumentException("SendKeysKeywordDelimError")
                End If
                Dim keyword As String = keys.Substring((num + 1), (num6 - (num + 1)))
                If Char.IsWhiteSpace(keys.Chars(num6)) Then
                    Do While ((num6 < length) AndAlso Char.IsWhiteSpace(keys.Chars(num6)))
                        num6 += 1
                    Loop
                    If (num6 >= length) Then
                        Throw New ArgumentException("SendKeysKeywordDelimError")
                    End If
                    If Char.IsDigit(keys.Chars(num6)) Then
                        Dim startIndex As Integer = num6
                        Do While ((num6 < length) AndAlso Char.IsDigit(keys.Chars(num6)))
                            num6 += 1
                        Loop
                        repeat = Integer.Parse(keys.Substring(startIndex,
                          (num6 - startIndex)), CultureInfo.InvariantCulture)
                    End If
                End If
                If (num6 >= length) Then
                    Throw New ArgumentException("SendKeysKeywordDelimError")
                End If
                If (keys.Chars(num6) <> "}"c) Then
                    Throw New ArgumentException("InvalidSendKeysRepeat")
                End If
                vk = SendKeysPlus.MatchKeyword(keyword)
                If (vk <> -1) Then
                    If ((haveKeys(0) = 0) AndAlso ((vk And &H10000) <> 0)) Then
                        SendKeysPlus.AddEvent(New SKEvent(&H100, &H10,
                          SendKeysPlus.fStartNewChar, hwnd, ScanCodeMode))
                        SendKeysPlus.fStartNewChar = False
                        haveKeys(0) = 10
                    End If
                    If ((haveKeys(1) = 0) AndAlso ((vk And &H20000) <> 0)) Then
                        SendKeysPlus.AddEvent(New SKEvent(&H100, &H11,
                          SendKeysPlus.fStartNewChar, hwnd, ScanCodeMode))
                        SendKeysPlus.fStartNewChar = False
                        haveKeys(1) = 10
                    End If
                    If ((haveKeys(2) = 0) AndAlso ((vk And &H40000) <> 0)) Then
                        SendKeysPlus.AddEvent(New SKEvent(&H100, &H12,
                           SendKeysPlus.fStartNewChar, hwnd, ScanCodeMode))
                        SendKeysPlus.fStartNewChar = False
                        haveKeys(2) = 10
                    End If
                    SendKeysPlus.AddMsgsForVK(vk, repeat, ((haveKeys(2) > 0) _
                       AndAlso (haveKeys(1) = 0)), hwnd, ScanCodeMode)
                    SendKeysPlus.CancelMods(haveKeys, 10, hwnd, ScanCodeMode)
                Else
                    If (keyword.Length <> 1) Then
                        Throw New ArgumentException("InvalidSendKeysKeyword" &
                         keys.Substring((num + 1), (num6 - (num + 1))))
                    End If
                    SendKeysPlus.fStartNewChar = SendKeysPlus.AddSimpleKey(keyword.Chars(0),
                      repeat, hwnd, haveKeys, SendKeysPlus.fStartNewChar, cGrp, ScanCodeMode)
                End If
                num = num6
                GoTo Label_04AB
Label_0333:
                SendKeysPlus.AddEvent(New SKEvent(&H100, &H10,
                  SendKeysPlus.fStartNewChar, hwnd, ScanCodeMode))
                SendKeysPlus.fStartNewChar = False
                haveKeys(0) = 10
                GoTo Label_04AB
Label_03C9:
                SendKeysPlus.AddEvent(New SKEvent(If((haveKeys(1) <> 0),
                   &H100, 260), &H12, SendKeysPlus.fStartNewChar, hwnd, ScanCodeMode))
                SendKeysPlus.fStartNewChar = False
                haveKeys(2) = 10
                GoTo Label_04AB
Label_0414:
                If (haveKeys(0) = 10) Then
                    haveKeys(0) = cGrp
                End If
                If (haveKeys(1) = 10) Then
                    haveKeys(1) = cGrp
                End If
                If (haveKeys(2) = 10) Then
                    haveKeys(2) = cGrp
                End If
                GoTo Label_04AB
Label_045A:
                SendKeysPlus.CancelMods(haveKeys, cGrp, hwnd, ScanCodeMode)
                cGrp -= 1
                If (cGrp = 0) Then
                    SendKeysPlus.fStartNewChar = True
                End If
Label_04AB:
                num += 1
            Loop
            If (cGrp <> 0) Then
                Throw New ArgumentException("SendKeysGroupDelimError")
            End If
            SendKeysPlus.CancelMods(haveKeys, 10, hwnd, ScanCodeMode)
        End Sub

        Private Shared Sub ResetKeyboardUsingSendInput(ByVal INPUTSize As Integer)
            If ((SendKeysPlus.capslockChanged OrElse SendKeysPlus.numlockChanged) _
             OrElse (SendKeysPlus.scrollLockChanged OrElse SendKeysPlus.kanaChanged)) Then
                Dim pInputs As INPUT() = New INPUT(2 - 1) {}
                pInputs(0).type = 1
                pInputs(0).ki.dwFlags = 0
                pInputs(1).type = 1
                pInputs(1).ki.dwFlags = 2
                If SendKeysPlus.capslockChanged Then
                    pInputs(0).ki.wVk = 20
                    pInputs(1).ki.wVk = 20
                    SendInput(2, pInputs, INPUTSize)
                End If
                If SendKeysPlus.numlockChanged Then
                    pInputs(0).ki.wVk = &H90
                    pInputs(1).ki.wVk = &H90
                    SendInput(2, pInputs, INPUTSize)
                End If
                If SendKeysPlus.scrollLockChanged Then
                    pInputs(0).ki.wVk = &H91
                    pInputs(1).ki.wVk = &H91
                    SendInput(2, pInputs, INPUTSize)
                End If
                If SendKeysPlus.kanaChanged Then
                    pInputs(0).ki.wVk = &H15
                    pInputs(1).ki.wVk = &H15
                    SendInput(2, pInputs, INPUTSize)
                End If
            End If
        End Sub

        ''' <summary>Sends keystrokes to the active application.</summary>
        ''' <param name="keys">The string of keystrokes to send. </param>
        ''' <exception cref="T:System.InvalidOperationException">There is not
        ''' an active application to send keystrokes to. </exception>
        ''' <exception cref="T:System.ArgumentException">keys do 
        '''    not represent valid keystrokes</exception>
        ''' <filterpriority>1</filterpriority>
        ''' <PermissionSet><IPermission
        '''   class="System.Security.Permissions.EnvironmentPermission, 
        '''   mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Unrestricted="true" />
        ''' <IPermission class="System.Security.Permissions.FileIOPermission, 
        '''    mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''    version="1" Unrestricted="true" />
        ''' <IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, 
        '''   Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Flags="UnmanagedCode, ControlEvidence" />
        ''' <IPermission class="System.Security.Permissions.UIPermission, mscorlib, 
        '''    Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''    version="1" Unrestricted="true" />
        ''' <IPermission class="System.Diagnostics.PerformanceCounterPermission, 
        '''   System, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Unrestricted="true" /></PermissionSet>
        Public Shared Sub Send(ByVal keys As String, ByVal ScanCodeMode As ScanCodeModes)
            SendKeysPlus.Send(keys, Nothing, False, ScanCodeMode)
        End Sub

        Private Shared Sub Send(ByVal keys As String, ByVal hWnd As IntPtr,
     ByVal wait As Boolean, ByVal ScanCodeMode As ScanCodeModes)

            '!!! IntSecurity.UnmanagedCode.Demand()
            If ((Not keys Is Nothing) AndAlso (keys.Length <> 0)) Then
                If (Not wait AndAlso Not Application.MessageLoop) Then
                    Throw New InvalidOperationException("SendKeysNoMessageLoop")
                End If
                Dim previousEvents As Queue = Nothing
                If ((Not SendKeysPlus.events Is Nothing) AndAlso (SendKeysPlus.events.Count <> 0)) Then
                    previousEvents = DirectCast(SendKeysPlus.events.Clone, Queue)
                End If
                SendKeysPlus.ParseKeys(keys, hWnd, ScanCodeMode)
                If (Not SendKeysPlus.events Is Nothing) Then
                    SendKeysPlus.LoadSendMethodFromConfig()
                    Dim keyboardState As Byte() = SendKeysPlus.MyGetKeyboardState
                    If (SendKeysPlus.sendMethod.Value <> SendMethodTypes.SendInput) Then
                        If (Not SendKeysPlus.hookSupported.HasValue AndAlso
                         (SendKeysPlus.sendMethod.Value = SendMethodTypes.Default)) Then
                            SendKeysPlus.TestHook()
                        End If
                        If ((SendKeysPlus.sendMethod.Value = SendMethodTypes.JournalHook) _
                            OrElse SendKeysPlus.hookSupported.Value) Then
                            SendKeysPlus.ClearKeyboardState()
                            SendKeysPlus.InstallHook()
                            SendKeysPlus.MySetKeyboardState(keyboardState)
                        End If
                    End If
                    If ((SendKeysPlus.sendMethod.Value = SendMethodTypes.SendInput) OrElse
                    ((SendKeysPlus.sendMethod.Value = SendMethodTypes.Default) _
                    AndAlso Not SendKeysPlus.hookSupported.Value)) Then
                        SendKeysPlus.SendInput(keyboardState, previousEvents)
                    End If
                    If wait Then
                        SendKeysPlus.Flush()
                    End If
                End If
            End If
        End Sub




        Private Shared Sub SendInput(ByVal oldKeyboardState As Byte(), ByVal previousEvents As Queue)
            Dim count As Integer
            SendKeysPlus.AddCancelModifiersForPreviousEvents(previousEvents)
            Dim pInputs As INPUT() = New INPUT(2 - 1) {}
            pInputs(0).type = 1
            pInputs(1).type = 1
            pInputs(1).ki.wVk = 0
            pInputs(1).ki.dwFlags = 6
            pInputs(0).ki.dwExtraInfo = IntPtr.Zero
            pInputs(0).ki.time = 0
            pInputs(1).ki.dwExtraInfo = IntPtr.Zero
            pInputs(1).ki.time = 0
            Dim cbSize As Integer = Marshal.SizeOf(GetType(INPUT))
            Dim num2 As UInt32 = 0
            SyncLock SendKeysPlus.events.SyncRoot
                Dim flag As Boolean = BlockInput(True)
                Try
                    count = SendKeysPlus.events.Count
                    SendKeysPlus.ClearGlobalKeys()
                    Dim i As Integer
                    For i = 0 To count - 1
                        Dim skEvent As SKEvent = DirectCast(SendKeysPlus.events.Dequeue, SKEvent)
                        pInputs(0).ki.dwFlags = 0
                        If (skEvent.wm = &H102) Then  'OEMScanKey Set use VkKey
                            pInputs(0).ki.wVk = 0
                            pInputs(0).ki.wScan = CShort(skEvent.paramL)
                            pInputs(0).ki.dwFlags = 4
                            pInputs(1).ki.wScan = CShort(skEvent.paramL)
                            num2 = (num2 + (SendInput(2, pInputs, cbSize) - 1))
                        Else
                            If ((skEvent.wm = &H101) OrElse (skEvent.wm = &H105)) Then
                                pInputs(0).ki.dwFlags = (pInputs(0).ki.dwFlags Or 2)
                            End If
                            If SendKeysPlus.IsExtendedKey(skEvent) Then
                                pInputs(0).ki.dwFlags = (pInputs(0).ki.dwFlags Or 1)
                            End If

                            pInputs(0).ki.wVk = CShort(skEvent.paramL)

                            If skEvent.ScanCodeMode <> ScanCodeModes.None Then
                                Dim myKey As Keys
                                myKey = System.Enum.Parse(GetType(ScanCodeModes), CInt(skEvent.paramL))
                                pInputs(0).ki.wScan = GetScanCode(myKey, skEvent.ScanCodeMode)
                            Else
                                pInputs(0).ki.wScan = 0
                            End If
                            num2 = (num2 + SendInput(1, pInputs, cbSize))
                            SendKeysPlus.CheckGlobalKeys(skEvent)
                        End If
                        Thread.Sleep(1)
                    Next i
                    SendKeysPlus.ResetKeyboardUsingSendInput(cbSize)
                Finally
                    SendKeysPlus.MySetKeyboardState(oldKeyboardState)
                    If flag Then
                        BlockInput(False)
                    End If
                End Try
            End SyncLock
            If (num2 <> count) Then
                Throw New Win32Exception
            End If
        End Sub

        Public Enum ScanCodeModes
            None = 0
            Make = 1
            Break = 2
        End Enum
        Private _ScanCodeMode As ScanCodeModes = ScanCodeModes.None
        Public Property ScanCodeMode() As ScanCodeModes
            Get
                Return _ScanCodeMode
            End Get
            Set(ByVal value As ScanCodeModes)
                _ScanCodeMode = value
            End Set
        End Property



        ''' <summary>Sends the given keys to the active application,
        '''   and then waits for the messages to be processed.</summary>
        ''' <param name="keys">The string of keystrokes to send. </param>
        ''' <filterpriority>1</filterpriority>
        ''' <PermissionSet><IPermission 
        '''    class="System.Security.Permissions.EnvironmentPermission, 
        '''           mscorlib, Version=2.0.3600.0, Culture=neutral, 
        '''           PublicKeyToken=b77a5c561934e089" version="1" 
        '''    Unrestricted="true" />
        ''' <IPermission 
        '''   class="System.Security.Permissions.FileIOPermission, 
        '''         mscorlib, Version=2.0.3600.0, Culture=neutral, 
        '''         PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Unrestricted="true" />
        ''' <IPermission 
        '''    class="System.Security.Permissions.SecurityPermission, mscorlib, 
        '''           Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''    version="1" Flags="UnmanagedCode, ControlEvidence" />
        ''' <IPermission 
        '''   class="System.Security.Permissions.UIPermission, mscorlib, 
        '''         Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Unrestricted="true" />
        ''' <IPermission class="System.Diagnostics.PerformanceCounterPermission, 
        '''       System, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" 
        '''   version="1" Unrestricted="true" />
        ''' </PermissionSet>
        Public Shared Sub SendWait(ByVal keys As String, ByVal ScanCodeMode As ScanCodeModes)
            SendKeysPlus.SendWait(keys, IntPtr.Zero, ScanCodeMode)
        End Sub

        Public Shared Sub SendWait(ByVal keys As String, ByVal control As Control, ByVal ScanCodeMode As ScanCodeModes)
            SendKeysPlus.Send(keys, control.Handle, True, ScanCodeMode)
        End Sub

        Public Shared Sub SendWait(ByVal keys As String, ByVal hWnd As IntPtr, ByVal ScanCodeMode As ScanCodeModes)
            SendKeysPlus.Send(keys, hWnd, True, ScanCodeMode)
        End Sub

        Private Shared Sub MySetKeyboardState(ByVal keystate As Byte())
            SetKeyboardState(keystate)
        End Sub

        Private Shared Sub TestHook()
            SendKeysPlus.hookSupported = False
            Try
                Dim pfnhook As CallBack = New CallBack(AddressOf SendKeysPlus.EmptyHookCallback)
                Dim handle As IntPtr = SetWindowsHookEx(1,
                 pfnhook, New HandleRef(Nothing, GetModuleHandle(Nothing)), 0)
                SendKeysPlus.hookSupported = New Nullable(Of Boolean)((handle <> IntPtr.Zero))
                If (handle <> IntPtr.Zero) Then
                    UnhookWindowsHookEx(New HandleRef(Nothing, handle))
                End If
            Catch
            End Try
        End Sub

        Private Shared Sub UninstallJournalingHook()
            If (SendKeysPlus.hhook <> IntPtr.Zero) Then
                SendKeysPlus.stopHook = False
                If (Not SendKeysPlus.events Is Nothing) Then
                    SendKeysPlus.events.Clear()
                End If
                UnhookWindowsHookEx(New HandleRef(Nothing, SendKeysPlus.hhook))
                SendKeysPlus.hhook = IntPtr.Zero
            End If
        End Sub


        ' Fields
        Private Const ALTKEYSCAN As Integer = &H400
        Private Shared capslockChanged As Boolean
        Private Const CTRLKEYSCAN As Integer = &H200
        Private Shared events As Queue
        Private Shared fStartNewChar As Boolean
        Private Const HAVEALT As Integer = 2
        Private Const HAVECTRL As Integer = 1
        Private Const HAVESHIFT As Integer = 0
        Private Shared hhook As IntPtr
        Private Shared hook As CallBack
        Private Shared hookSupported As Nullable(Of Boolean) = Nothing
        Private Shared kanaChanged As Boolean
        Private Shared keywords As KeywordVk() =
          New KeywordVk() {New KeywordVk("ENTER", 13),
          New KeywordVk("TAB", 9), New KeywordVk("ESC", &H1B),
          New KeywordVk("ESCAPE", &H1B), New KeywordVk("HOME", &H24),
          New KeywordVk("END", &H23), New KeywordVk("LEFT", &H25),
          New KeywordVk("RIGHT", &H27), New KeywordVk("UP", &H26),
          New KeywordVk("DOWN", 40), New KeywordVk("PGUP", &H21),
          New KeywordVk("PGDN", &H22), New KeywordVk("NUMLOCK", &H90),
          New KeywordVk("SCROLLLOCK", &H91), New KeywordVk("PRTSC", &H2C),
          New KeywordVk("BREAK", 3), New KeywordVk("BACKSPACE", 8),
          New KeywordVk("BKSP", 8), New KeywordVk("BS", 8),
          New KeywordVk("CLEAR", 12), New KeywordVk("CAPSLOCK", 20),
          New KeywordVk("INS", &H2D), New KeywordVk("INSERT", &H2D),
          New KeywordVk("DEL", &H2E), New KeywordVk("DELETE", &H2E),
          New KeywordVk("HELP", &H2F), New KeywordVk("F1", &H70),
          New KeywordVk("F2", &H71), New KeywordVk("F3", &H72),
          New KeywordVk("F4", &H73), New KeywordVk("F5", &H74),
          New KeywordVk("F6", &H75), New KeywordVk("F7", &H76),
          New KeywordVk("F8", &H77), New KeywordVk("F9", 120),
          New KeywordVk("F10", &H79), New KeywordVk("F11", &H7A),
          New KeywordVk("F12", &H7B), New KeywordVk("F13", &H7C),
          New KeywordVk("F14", &H7D), New KeywordVk("F15", &H7E),
          New KeywordVk("F16", &H7F), New KeywordVk("MULTIPLY", &H6A),
          New KeywordVk("ADD", &H6B), New KeywordVk("SUBTRACT", &H6D),
          New KeywordVk("DIVIDE", &H6F), New KeywordVk("+", &H6B),
          New KeywordVk("%", &H10035), New KeywordVk("^", &H10036)}
        Private Shared messageWindow As SKWindow
        Private Shared numlockChanged As Boolean
        Private Shared scrollLockChanged As Boolean
        Private Shared sendMethod As Nullable(Of SendMethodTypes) = Nothing
        Private Const SHIFTKEYSCAN As Integer = &H100
        Private Shared stopHook As Boolean
        Private Const UNKNOWN_GROUPING As Integer = 10

        ' Nested Types
        Private Class KeywordVk
            ' Methods
            Public Sub New(ByVal key As String, ByVal v As Integer)
                Me.keyword = key
                Me.vk = v
            End Sub


            ' Fields
            Friend keyword As String
            Friend vk As Integer
        End Class

        Private Class SendKeysHookProc
            ' Methods
            Public Overridable Function Callback(ByVal code As Integer,
             ByVal wparam As IntPtr, ByVal lparam As IntPtr) As IntPtr
                Dim [Structure] As EVENTMSG =
                   DirectCast(Marshal.PtrToStructure(lparam,
                   GetType(EVENTMSG)), EVENTMSG)
                If (GetAsyncKeyState(&H13) <> 0) Then
                    SendKeysPlus.stopHook = True
                End If
                Select Case code
                    Case 1
                        Me.gotNextEvent = True
                        Dim event2 As SKEvent = DirectCast(SendKeysPlus.events.Peek, SKEvent)
                        [Structure].message = event2.wm
                        [Structure].paramL = event2.paramL
                        [Structure].paramH = event2.paramH
                        [Structure].hwnd = event2.hwnd
                        [Structure].time = GetTickCount()
                        Marshal.StructureToPtr([Structure], lparam, True)
                        Exit Select
                    Case 2
                        If Me.gotNextEvent Then
                            If ((Not SendKeysPlus.events Is Nothing) _
                               AndAlso (SendKeysPlus.events.Count > 0)) Then
                                SendKeysPlus.events.Dequeue()
                            End If
                            SendKeysPlus.stopHook =
                              ((SendKeysPlus.events Is Nothing) _
                              OrElse (SendKeysPlus.events.Count = 0))
                        End If
                        Exit Select
                    Case Else
                        If (code < 0) Then
                            CallNextHookEx(New HandleRef(Nothing, SendKeysPlus.hhook),
                        code, wparam, lparam)
                        End If
                        Exit Select
                End Select
                If SendKeysPlus.stopHook Then
                    SendKeysPlus.UninstallJournalingHook()
                    Me.gotNextEvent = False
                End If
                Return IntPtr.Zero
            End Function


            ' Fields
            Private gotNextEvent As Boolean
        End Class

        Public Enum SendMethodTypes
            ' Fields
            [Default] = 1
            JournalHook = 2
            SendInput = 3
        End Enum

        Private Class SKEvent
            ' Methods
            Public Sub New(ByVal a As Integer, ByVal b As Integer,
             ByVal c As Boolean, ByVal hwnd As IntPtr,
             ByVal ScanCodeMode As ScanCodeModes)
                Me.wm = a
                Me.paramL = b
                Me.paramH = If(c, 1, 0)
                Me.hwnd = hwnd
                Me.sc = GetScanCode(b, ScanCodeMode)
                Me.ScanCodeMode = ScanCodeMode
            End Sub

            Public Sub New(ByVal a As Integer, ByVal b As Integer,
             ByVal c As Integer, ByVal hwnd As IntPtr,
             ByVal ScanCodeMode As ScanCodeModes)
                Me.wm = a
                Me.paramL = b
                Me.paramH = c
                Me.hwnd = hwnd
                Me.sc = GetScanCode(b, ScanCodeMode)
                Me.ScanCodeMode = ScanCodeMode
            End Sub

            ' Fields
            Friend hwnd As IntPtr
            Friend paramH As Integer
            Friend paramL As Integer
            Friend wm As Integer
            Friend sc As Integer
            Friend ScanCodeMode As ScanCodeModes
        End Class

        Private Class SKWindow
            Inherits Control
            ' Methods
            Public Sub New()

                'MyBase.SetState(&H80000, True)
                'MyBase.SetState2(8, False)
                MyBase.SetBounds(-1, -1, 0, 0)
                MyBase.Visible = False
            End Sub

            Protected Overrides Sub WndProc(ByRef m As Message)
                If (m.Msg = &H4B) Then
                    Try
                        SendKeysPlus.JournalCancel()
                    Catch
                    End Try
                End If
            End Sub

        End Class

    End Class
End Namespace