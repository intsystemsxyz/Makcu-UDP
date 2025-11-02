Imports System.IO.Ports
Imports System.Threading

Namespace MouseMovementLibraries.MakcuSupport

    Public Enum MakcuMouseButton
        Left = 0
        Right = 1
        Middle = 2
        Mouse4 = 3
        Mouse5 = 4
    End Enum

    Public Class MakcuMouse
        Implements System.IDisposable

        Private Shared ReadOnly BaudChangeCommand() As System.Byte = New System.Byte() {&HDE, &HAD, &H5, &H0, &HA5, &H0, &H9, &H3D, &H0}

        Private _serialPort As System.IO.Ports.SerialPort = Nothing
        Private ReadOnly _debugLogging As System.Boolean
        Private ReadOnly _sendInitCommands As System.Boolean

        Private ReadOnly _serialLock As System.Object = New System.Object()
        Private _isInitializedAndConnected As System.Boolean = False
        Private _listenerThread As System.Threading.Thread = Nothing
        Private _stopListenerEvent As System.Threading.ManualResetEventSlim = New System.Threading.ManualResetEventSlim(False)
        Private _pauseListener As System.Boolean = False

        Private ReadOnly _buttonStates As System.Collections.Generic.Dictionary(Of MouseMovementLibraries.MakcuSupport.MakcuMouseButton, System.Boolean) = New System.Collections.Generic.Dictionary(Of MouseMovementLibraries.MakcuSupport.MakcuMouseButton, System.Boolean)()

        Public Event ButtonStateChanged As System.Action(Of MouseMovementLibraries.MakcuSupport.MakcuMouseButton, System.Boolean)

        Public Property PortName As System.String

        Public ReadOnly Property BaudRate() As System.Int32
            Get
                If _serialPort IsNot Nothing Then
                    Return _serialPort.BaudRate
                Else
                    If _isInitializedAndConnected Then
                        Return 4000000
                    Else
                        Return 115200
                    End If
                End If
            End Get
        End Property

        Public ReadOnly Property IsInitializedAndConnected() As System.Boolean
            Get
                Return _isInitializedAndConnected AndAlso _serialPort IsNot Nothing AndAlso _serialPort.IsOpen
            End Get
        End Property

        Public Sub New(Optional ByVal debugLogging As System.Boolean = False, Optional ByVal sendInitCommands As System.Boolean = True)
            _debugLogging = debugLogging
            _sendInitCommands = sendInitCommands

            For Each btn As MouseMovementLibraries.MakcuSupport.MakcuMouseButton In System.[Enum].GetValues(GetType(MouseMovementLibraries.MakcuSupport.MakcuMouseButton))
                _buttonStates(btn) = False
            Next
        End Sub

        Private Sub Log(ByVal message As System.String)
            If _debugLogging Then
                System.Console.WriteLine("[MakcuMouse " & System.DateTime.Now.ToString("HH:mm:ss") & "] " & message)
            End If
        End Sub

        Private Function FindComPortInternal() As System.String
            Log("Searching for COM port for Makcu device using SerialPort.GetPortNames() and connection test...")
            Dim availablePorts() As System.String = Nothing

            Try
                availablePorts = System.IO.Ports.SerialPort.GetPortNames()
            Catch ex As System.Exception
                Log("Error getting COM port list: " & ex.Message & ". Cannot continue search.")
                Return Nothing
            End Try

            If availablePorts Is Nothing OrElse Not availablePorts.Any() Then
                Log("No COM ports available according to SerialPort.GetPortNames().")
                Return Nothing
            End If

            Log("Available COM ports: " & System.String.Join(", ", availablePorts) & ". Testing each one...")

            For Each portName As System.String In availablePorts
                Log("Testing port: " & portName)
                Dim testPort As System.IO.Ports.SerialPort = Nothing

                Try
                    testPort = New System.IO.Ports.SerialPort(portName, 115200)
                    testPort.ReadTimeout = 250
                    testPort.WriteTimeout = 500
                    testPort.DtrEnable = True
                    testPort.RtsEnable = True
                    testPort.Open()
                    Log("Port " & portName & " opened at 115200 baud.")

                    Log("Sending command to change baudrate to 4M on " & portName & ".")
                    testPort.Write(BaudChangeCommand, 0, BaudChangeCommand.Length)
                    testPort.BaseStream.Flush()
                    System.Threading.Thread.Sleep(150)

                    testPort.Close()
                    testPort.BaudRate = 4000000
                    testPort.Open()
                    Log("Port " & portName & " reopened at 4000000 baud.")

                    testPort.DiscardInBuffer()
                    testPort.Write("km.version()" & vbCrLf)
                    Log("Command 'km.version()' sent to " & portName & ". Waiting for response...")

                    Dim response As System.String = System.String.Empty
                    Try
                        response = testPort.ReadLine().Trim()
                    Catch tex As System.TimeoutException
                        Log("Timeout waiting for 'km.version()' response on " & portName & ". Probably not the Makcu device.")
                        testPort.Close()
                        Continue For
                    End Try

                    Log("Response from 'km.version()' on " & portName & ": '" & response & "'")

                    Dim isMakcu As System.Boolean = False
                    If Not System.String.IsNullOrEmpty(response) Then
                        If response.Contains("KMBOX") OrElse response.Contains("Makcu") OrElse response.Contains("MAKCU") OrElse response.StartsWith("v", System.StringComparison.Ordinal) Then
                            isMakcu = True
                        Else
                            Dim ch As System.Char = response.FirstOrDefault()
                            If System.Char.IsDigit(ch) Then
                                isMakcu = True
                            End If
                        End If
                    End If

                    If isMakcu Then
                        Log("Makcu device found on port: " & portName & "! Version response: '" & response & "'")
                        testPort.Close()
                        Return portName
                    Else
                        Log("Response from 'km.version()' on " & portName & " does not seem to be from a Makcu device. Response: '" & response & "'")
                    End If

                    testPort.Close()

                Catch exUA As System.UnauthorizedAccessException
                    Log("Unauthorized access error on " & portName & ": " & exUA.Message & " (Port might be in use).")
                Catch exTO As System.TimeoutException
                    Log("Timeout on " & portName & " (possibly during baud rate change or write): " & exTO.Message)
                Catch ex As System.Exception
                    Log("Failed to test port " & portName & ": " & ex.[GetType]().Name & " - " & ex.Message)
                Finally
                    If testPort IsNot Nothing AndAlso testPort.IsOpen Then
                        testPort.Close()
                    End If
                    If testPort IsNot Nothing Then
                        testPort.Dispose()
                    End If
                End Try
            Next

            Log("Makcu device not found on any available COM ports after connection test.")
            Return Nothing
        End Function

        Private Function OpenSerialPort(ByVal portNameToOpen As System.String, ByVal baudRate As System.Int32) As System.Boolean
            Try
                If _serialPort IsNot Nothing Then
                    _serialPort.Close()
                    _serialPort.Dispose()
                    _serialPort = Nothing
                End If

                Log("Attempting to open " & portNameToOpen & " at " & baudRate.ToString() & " baud.")
                _serialPort = New System.IO.Ports.SerialPort(portNameToOpen, baudRate)
                _serialPort.ReadTimeout = 100
                _serialPort.WriteTimeout = 500
                _serialPort.DtrEnable = True
                _serialPort.RtsEnable = True
                _serialPort.Open()
                Me.PortName = portNameToOpen
                Return True
            Catch ex As System.Exception
                Log("Failed to open " & portNameToOpen & " at " & baudRate.ToString() & " baud. Error: " & ex.Message)
                If _serialPort IsNot Nothing Then
                    _serialPort.Dispose()
                    _serialPort = Nothing
                End If
                Return False
            End Try
        End Function

        Private Function ChangeBaudRateTo4M() As System.Boolean
            If _serialPort Is Nothing OrElse Not _serialPort.IsOpen Then
                Log("ChangeBaudRateTo4M: _serialPort is null or not open.")
                Return False
            End If

            Log("Sending command to change baudrate to 4M.")
            Try
                _serialPort.Write(BaudChangeCommand, 0, BaudChangeCommand.Length)
                _serialPort.BaseStream.Flush()
                System.Threading.Thread.Sleep(150)

                Dim currentPortName As System.String = _serialPort.PortName
                _serialPort.Close()

                _serialPort.BaudRate = 4000000
                _serialPort.Open()

                Log("Successfully changed to 4M baud.")
                Return True
            Catch ex As System.Exception
                Log("Error changing baudrate to 4M: " & ex.Message)
                If _serialPort IsNot Nothing Then
                    Try
                        _serialPort.Close()
                    Catch
                    End Try
                End If
                Return False
            End Try
        End Function

        Public Function Init() As System.Boolean
            SyncLock _serialLock
                If _isInitializedAndConnected Then
                    Log("MakcuMouse is already initialized.")
                    Return True
                End If

                Dim portToUse As System.String = FindComPortInternal()
                If System.String.IsNullOrEmpty(portToUse) Then
                    Log("Init failed: Could not find a valid COM port for Makcu.")
                    Return False
                End If

                If Not OpenSerialPort(portToUse, 115200) Then
                    Log("Init failed: Could not open port " & portToUse & " at 115200 baud.")
                    Return False
                End If

                If Not ChangeBaudRateTo4M() Then
                    Log("Init failed: Could not change to 4M baud.")
                    Close()
                    Return False
                End If

                _isInitializedAndConnected = True
                Me.PortName = portToUse

                If _sendInitCommands Then
                    Dim tmp As System.String = Nothing
                    If Not SendCommandInternal("km.buttons(1)", False, tmp) Then
                        Log("Init failed: Could not send 'km.buttons(1)'.")
                        Close()
                        Return False
                    End If

                    _stopListenerEvent.Reset()
                    _listenerThread = New System.Threading.Thread(Sub() ListenForButtonEvents(_debugLogging))
                    _listenerThread.IsBackground = True
                    _listenerThread.Name = "MakcuButtonListener"
                    _listenerThread.Start()
                End If

                Log("MakcuMouse initialized successfully on port " & Me.PortName & " at " & Me.BaudRate.ToString() & " baud.")
                Return True
            End SyncLock
        End Function

        Private Function SendCommandInternal(ByVal command As System.String,
                                             ByVal expectResponse As System.Boolean,
                                             ByRef responseText As System.String,
                                             Optional ByVal responseTimeoutMs As System.Int32 = 200) As System.Boolean
            responseText = Nothing

            If Not _isInitializedAndConnected OrElse _serialPort Is Nothing OrElse Not _serialPort.IsOpen Then
                Log("Error in SendCommand: Connection not open or initialized. Command: " & command)
                Return False
            End If

            SyncLock _serialLock
                If Not _serialPort.IsOpen Then
                    Log("Error in SendCommand (lock): Port not open. Command: " & command)
                    Return False
                End If

                _pauseListener = True
                Dim originalReadTimeout As System.Int32 = _serialPort.ReadTimeout

                Try
                    If expectResponse Then
                        _serialPort.ReadTimeout = responseTimeoutMs
                        _serialPort.DiscardInBuffer()
                    End If

                    Log("Sending: " & command)
                    _serialPort.Write(command & vbCrLf)

                    If expectResponse Then
                        Dim stopwatch As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
                        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
                        Dim commandTrimmed As System.String = command.Trim()
                        Dim firstLineIsEcho As System.Boolean = True

                        While stopwatch.ElapsedMilliseconds < responseTimeoutMs
                            Try
                                Dim line As System.String = _serialPort.ReadLine().Trim()
                                If Not System.String.IsNullOrEmpty(line) Then
                                    Log("Received (raw): " & line)
                                    If line.Equals(commandTrimmed, System.StringComparison.OrdinalIgnoreCase) AndAlso firstLineIsEcho Then
                                        firstLineIsEcho = False
                                        Continue While
                                    End If
                                    If line.Equals("OK", System.StringComparison.OrdinalIgnoreCase) Then
                                        If sb.Length > 0 Then
                                            ' swallow trailing OK
                                            Continue While
                                        End If
                                    End If
                                    If sb.Length > 0 Then
                                        sb.Append(vbLf)
                                    End If
                                    sb.Append(line)
                                End If

                                If _serialPort.BytesToRead = 0 AndAlso sb.Length > 0 Then
                                    System.Threading.Thread.Sleep(15)
                                    If _serialPort.BytesToRead = 0 Then
                                        Exit While
                                    End If
                                End If
                            Catch tex As System.TimeoutException
                                Log("ReadLine() timeout in SendCommandInternal for '" & command & "'.")
                                Exit While
                            Catch ioe As System.InvalidOperationException
                                Log("InvalidOperationException in ReadLine (SendCommandInternal): " & ioe.Message)
                                _isInitializedAndConnected = False
                                Close()
                                Return False
                            Catch ex As System.Exception
                                Log("Exception in ReadLine (SendCommandInternal): " & ex.Message)
                                Exit While
                            End Try
                        End While

                        stopwatch.Stop()
                        responseText = sb.ToString().Trim()
                        Log("Final response for '" & command & "': '" & responseText & "' (Time: " & stopwatch.ElapsedMilliseconds.ToString() & "ms)")
                        Return True
                    End If

                    Return True

                Catch tex As System.TimeoutException
                    Log("General timeout in SendCommandInternal for '" & command & "': " & tex.Message)
                    Return False

                Catch ioe As System.InvalidOperationException
                    Log("General InvalidOperationException in SendCommandInternal (possibly port closed): " & ioe.Message)
                    _isInitializedAndConnected = False
                    Close()
                    Return False

                Catch ex As System.Exception
                    Log("General exception in SendCommandInternal for '" & command & "': " & ex.Message)
                    Return False

                Finally
                    _pauseListener = False
                    If expectResponse Then
                        _serialPort.ReadTimeout = originalReadTimeout
                    End If
                End Try
            End SyncLock
        End Function

        Private Sub ListenForButtonEvents(ByVal debug As System.Boolean)
            Log("Listener thread started. (PACKET PARSING MODE)")

            Dim lastMask As System.Byte = CByte(&H0)

            Dim buttonMap As System.Collections.Generic.Dictionary(Of System.Int32, MouseMovementLibraries.MakcuSupport.MakcuMouseButton) =
                New System.Collections.Generic.Dictionary(Of System.Int32, MouseMovementLibraries.MakcuSupport.MakcuMouseButton)() From {
                    {0, MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left},
                    {1, MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right},
                    {2, MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle},
                    {3, MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse4},
                    {4, MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse5}
                }

            Dim packetHeader() As System.Byte = New System.Byte() {&H6B, &H6D, &H2E}
            Dim currentHeaderIndex As System.Int32 = 0

            While Not _stopListenerEvent.IsSet
                If Not Me.IsInitializedAndConnected OrElse _pauseListener OrElse _serialPort Is Nothing OrElse Not _serialPort.IsOpen Then
                    System.Threading.Thread.Sleep(20)
                    currentHeaderIndex = 0
                    Continue While
                End If

                Try
                    If _serialPort.BytesToRead > 0 Then
                        Dim byteRead As System.Byte = System.Convert.ToByte(_serialPort.ReadByte())

                        If byteRead = packetHeader(currentHeaderIndex) Then
                            currentHeaderIndex += 1

                            If currentHeaderIndex = packetHeader.Length Then
                                If debug Then
                                    Log("Listener: Packet header " & System.String.Join(",", packetHeader.Select(Function(b As System.Byte) b.ToString("X2"))) & " detected!")
                                End If

                                If _serialPort.BytesToRead > 0 Then
                                    Dim currentMask As System.Byte = System.Convert.ToByte(_serialPort.ReadByte())
                                    If debug Then
                                        Log("Listener: Potential button mask byte: 0x" & currentMask.ToString("X2"))
                                    End If

                                    If currentMask <= CByte(&H1F) Then
                                        If currentMask <> lastMask Then
                                            If debug Then
                                                Log("Listener: Processing button mask. New: 0x" & currentMask.ToString("X2") & ", Prev: 0x" & lastMask.ToString("X2"))
                                            End If

                                            Dim changedBits As System.Byte = System.Convert.ToByte(currentMask Xor lastMask)

                                            SyncLock _buttonStates
                                                For Each pair As System.Collections.Generic.KeyValuePair(Of System.Int32, MouseMovementLibraries.MakcuSupport.MakcuMouseButton) In buttonMap
                                                    If (changedBits And CByte(1 << pair.Key)) <> 0 Then
                                                        Dim isPressed As System.Boolean = (currentMask And CByte(1 << pair.Key)) <> 0
                                                        _buttonStates(pair.Value) = isPressed
                                                        If debug Then
                                                            Log("Listener: ---> EVENT: Button: " & pair.Value.ToString() & ", IsPressed: " & isPressed.ToString())
                                                        End If
                                                        Try
                                                            RaiseEvent ButtonStateChanged(pair.Value, isPressed)
                                                        Catch ex As System.Exception
                                                            Log("Exception in ButtonStateChanged handler: " & ex.Message)
                                                        End Try
                                                    End If
                                                Next
                                            End SyncLock

                                            lastMask = currentMask

                                            If debug Then
                                                Dim pressedButtons() As System.String = _buttonStates.Where(Function(kvp) kvp.Value).Select(Function(kvp) kvp.Key.ToString()).ToArray()
                                                Log("Listener: Button states updated. Mask: 0x" & currentMask.ToString("X2") & " -> " & If(pressedButtons.Any(), System.String.Join(", ", pressedButtons), "None"))
                                            End If
                                        End If
                                    End If

                                    Dim expectedTailBytes As System.Int32 = 2
                                    For i As System.Int32 = 0 To expectedTailBytes - 1
                                        If _serialPort.BytesToRead > 0 Then
                                            Dim consumedByte As System.Byte = System.Convert.ToByte(_serialPort.ReadByte())
                                            If debug Then
                                                Log("Listener: Consumed tail byte " & (i + 1).ToString() & ": 0x" & consumedByte.ToString("X2"))
                                            End If
                                        Else
                                            If debug Then
                                                Log("Listener: Expected tail byte " & (i + 1).ToString() & " but no data. Packet might be short.")
                                            End If
                                            Exit For
                                        End If
                                    Next
                                Else
                                    If debug Then
                                        Log("Listener: Header found, but no data for button mask byte. Packet might be short.")
                                    End If
                                End If

                                currentHeaderIndex = 0
                            End If
                        Else
                            If currentHeaderIndex > 0 AndAlso debug Then
                                Log("Listener: Byte 0x" & byteRead.ToString("X2") & " broke header sequence at index " & currentHeaderIndex.ToString() & ". Resetting search.")
                            End If
                            currentHeaderIndex = 0
                            If byteRead = packetHeader(0) Then
                                currentHeaderIndex = 1
                            End If
                        End If
                    Else
                        System.Threading.Thread.Sleep(1)
                    End If

                Catch tex As System.TimeoutException
                    If debug Then
                        Log("Listener: TimeoutException during serial read.")
                    End If
                    currentHeaderIndex = 0

                Catch ioe As System.InvalidOperationException
                    Log("Listener: InvalidOperationException (port probably closed): " & ioe.Message & ". Stopping listener.")
                    _isInitializedAndConnected = False
                    currentHeaderIndex = 0
                    Exit While

                Catch ex As System.Exception
                    Log("Error in listener: " & ex.[GetType]().Name & " - " & ex.Message)
                    currentHeaderIndex = 0
                    System.Threading.Thread.Sleep(100)
                End Try
            End While

            Log("Listener thread terminated.")
        End Sub

        Public Function Press(ByVal button As MouseMovementLibraries.MakcuSupport.MakcuMouseButton) As System.Boolean
            Dim dummy As System.String = Nothing
            Return SendCommandInternal("km." & GetButtonString(button) & "(1)", False, dummy)
        End Function

        Public Function Release(ByVal button As MouseMovementLibraries.MakcuSupport.MakcuMouseButton) As System.Boolean
            Dim dummy As System.String = Nothing
            Return SendCommandInternal("km." & GetButtonString(button) & "(0)", False, dummy)
        End Function

        Public Function Move(ByVal x As System.Int32, ByVal y As System.Int32) As System.Boolean
            Dim dummy As System.String = Nothing
            Return SendCommandInternal("km.move(" & x.ToString() & "," & y.ToString() & ")", False, dummy)
        End Function


        Public Function MoveSmooth(ByVal x As System.Int32, ByVal y As System.Int32, ByVal segments As System.Int32) As System.Boolean
            Dim dummy As System.String = Nothing
            Return SendCommandInternal("km.move(" & x.ToString() & "," & y.ToString() & "," & segments.ToString() & ")", False, dummy)
        End Function

        Public Function MoveBezier(ByVal x As System.Int32, ByVal y As System.Int32, ByVal segments As System.Int32, ByVal ctrlX As System.Int32, ByVal ctrlY As System.Int32) As System.Boolean
            Dim dummy As System.String = Nothing
            Return SendCommandInternal("km.move(" & x.ToString() & "," & y.ToString() & "," & segments.ToString() & "," & ctrlX.ToString() & "," & ctrlY.ToString() & ")", False, dummy)
        End Function

        Public Function Scroll(ByVal delta As System.Int32) As System.Boolean
            Dim dummy As System.String = Nothing
            Return SendCommandInternal("km.wheel(" & delta.ToString() & ")", False, dummy)
        End Function

        Public Function GetKmVersion() As System.String
            Dim response As System.String = Nothing
            Dim ok As System.Boolean = SendCommandInternal("km.version()", True, response, 500)
            If ok Then
                Return response
            Else
                Return Nothing
            End If
        End Function

        Private Function GetButtonString(ByVal button As MouseMovementLibraries.MakcuSupport.MakcuMouseButton) As System.String
            Select Case button
                Case MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left : Return "left"
                Case MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right : Return "right"
                Case MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle : Return "middle"
                Case Else
                    Throw New System.ArgumentException("Button " & button.ToString() & " not supported for direct press/release actions (left/right/middle).")
            End Select
        End Function

        Private Function GetButtonLockString(ByVal button As MouseMovementLibraries.MakcuSupport.MakcuMouseButton) As System.String
            Select Case button
                Case MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left : Return "ml"
                Case MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right : Return "mr"
                Case MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle : Return "mm"
                Case MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse4 : Return "ms1"
                Case MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse5 : Return "ms2"
                Case Else
                    Throw New System.ArgumentException("Button not supported for lock/catch: " & button.ToString())
            End Select
        End Function

        Public Function GetCurrentButtonStates() As System.Collections.Generic.Dictionary(Of MouseMovementLibraries.MakcuSupport.MakcuMouseButton, System.Boolean)
            SyncLock _buttonStates
                Return New System.Collections.Generic.Dictionary(Of MouseMovementLibraries.MakcuSupport.MakcuMouseButton, System.Boolean)(_buttonStates)
            End SyncLock
        End Function

        Public Function GetCurrentButtonMask() As System.Int32
            Dim mask As System.Int32 = 0
            SyncLock _buttonStates
                Dim l As System.Boolean = False
                Dim r As System.Boolean = False
                Dim m As System.Boolean = False
                Dim m4 As System.Boolean = False
                Dim m5 As System.Boolean = False

                If _buttonStates.TryGetValue(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left, l) AndAlso l Then
                    mask = mask Or (1 << CInt(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left))
                End If
                If _buttonStates.TryGetValue(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right, r) AndAlso r Then
                    mask = mask Or (1 << CInt(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right))
                End If
                If _buttonStates.TryGetValue(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle, m) AndAlso m Then
                    mask = mask Or (1 << CInt(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle))
                End If
                If _buttonStates.TryGetValue(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse4, m4) AndAlso m4 Then
                    mask = mask Or (1 << CInt(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse4))
                End If
                If _buttonStates.TryGetValue(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse5, m5) AndAlso m5 Then
                    mask = mask Or (1 << CInt(MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse5))
                End If
            End SyncLock
            Return mask
        End Function

        Private _disposedValue As System.Boolean = False

        Protected Overridable Sub Dispose(ByVal disposing As System.Boolean)
            If Not _disposedValue Then
                If disposing Then
                    Me.Close()
                    If _stopListenerEvent IsNot Nothing Then
                        _stopListenerEvent.Dispose()
                        _stopListenerEvent = Nothing
                    End If
                End If
                _disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            Me.Dispose(True)
            System.GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Me.Dispose(False)
            MyBase.Finalize()
        End Sub

        Public Sub Close()
            Me.Log("Closing MakcuMouse connection...")
            _isInitializedAndConnected = False

            If _listenerThread IsNot Nothing Then
                _stopListenerEvent.Set()
                If _listenerThread.IsAlive Then
                    If Not _listenerThread.Join(System.TimeSpan.FromSeconds(1)) Then
                        Me.Log("Listener thread did not terminate in time.")
                    End If
                End If
                _listenerThread = Nothing
            End If

            SyncLock _serialLock
                If _serialPort IsNot Nothing Then
                    If _serialPort.IsOpen Then
                        Try
                            _serialPort.Close()
                        Catch ex As System.Exception
                            Me.Log("Error closing port: " & ex.Message)
                        End Try
                    End If
                    _serialPort.Dispose()
                    _serialPort = Nothing
                End If
            End SyncLock

            Me.PortName = Nothing
            Me.Log("MakcuMouse connection closed.")
        End Sub






        ' RapidFire Variablen - ALLES PUBLIC
        Public RapidFireEnabled As Boolean = False
        Public RapidFireThread As Thread = Nothing
        Public RapidFireStopEvent As ManualResetEventSlim = New ManualResetEventSlim(False)
        Public RapidFireSpeedFactor As Integer = 3 ' 1 = langsam, 2 = mittel, 3 = schnell
        Public LeftButtonPressedFromRapidFire As Boolean = False

        ' Human-like Click Patterns
        Public HumanClickPatterns As Dictionary(Of String, Integer()) = New Dictionary(Of String, Integer()) From {
            {"slow", New Integer() {80, 90, 75, 95, 85, 70, 88, 78, 92, 82}},
            {"medium", New Integer() {45, 55, 40, 60, 50, 35, 58, 42, 52, 38}},
            {"fast", New Integer() {20, 25, 18, 28, 22, 16, 26, 19, 24, 17}},
            {"burst", New Integer() {15, 15, 15, 50, 15, 15, 15, 50, 15, 15}}
        }

        Public CurrentPattern As String = "fast"
        Public PatternIndex As Integer = 0
        Public LastClickTime As Long = 0

        ' Public Properties für einfachen Zugriff
        Public ReadOnly Property AvailablePatterns() As String()
            Get
                Return HumanClickPatterns.Keys.ToArray()
            End Get
        End Property

        Public ReadOnly Property CurrentClicksPerSecond() As Integer
            Get
                If HumanClickPatterns.ContainsKey(CurrentPattern) Then
                    Dim pattern = HumanClickPatterns(CurrentPattern)
                    Dim avgInterval = pattern.Average()
                    Dim speedMultiplier As Double = GetSpeedMultiplier()
                    Return CInt(1000.0 / (avgInterval / speedMultiplier))
                End If
                Return 0
            End Get
        End Property

        ' Public Methoden
        Public Function GetSpeedMultiplier() As Double
            Select Case RapidFireSpeedFactor
                Case 1 : Return 0.7 ' Langsam
                Case 2 : Return 1.0 ' Mittel
                Case 3 : Return 1.5 ' Schnell
                Case Else : Return 1.0
            End Select
        End Function

        Public Function GetHumanizedInterval() As Integer
            If HumanClickPatterns.ContainsKey(CurrentPattern) Then
                Dim pattern = HumanClickPatterns(CurrentPattern)

                ' Nächsten Wert aus dem Pattern nehmen
                Dim baseInterval As Integer = pattern(PatternIndex)
                PatternIndex = (PatternIndex + 1) Mod pattern.Length

                ' Leichte Zufallsabweichung für menschliches Feeling (±10%)
                Dim random As New Random()
                Dim variation As Double = random.NextDouble() * 0.2 - 0.1 ' -10% bis +10%
                Dim variedInterval As Integer = CInt(baseInterval * (1.0 + variation))

                ' Geschwindigkeitsfaktor anwenden
                Dim speedMultiplier As Double = GetSpeedMultiplier()
                Dim finalInterval As Integer = CInt(variedInterval / speedMultiplier)

                ' Minimum 8ms sicherstellen
                Return Math.Max(8, finalInterval)
            End If

            Return 50 ' Fallback
        End Function

        Public Function GetHumanizedClickDuration() As Integer
            ' Klick-Länge basierend auf Geschwindigkeitsfaktor
            Dim baseDuration As Integer = 12
            Dim speedMultiplier As Double = GetSpeedMultiplier()
            Dim variedDuration As Integer = CInt(baseDuration / speedMultiplier)
            Return Math.Max(3, Math.Min(20, variedDuration)) ' 3-20ms range
        End Function

        ' RapidFire Steuerung - PUBLIC
        Public Sub StartRapidFire()
            If RapidFireThread IsNot Nothing AndAlso RapidFireThread.IsAlive Then
                Return
            End If

            RapidFireStopEvent.Reset()
            PatternIndex = 0
            RapidFireThread = New Thread(AddressOf RapidFireWorker)
            RapidFireThread.IsBackground = True
            RapidFireThread.Name = "MakcuRapidFire"
            RapidFireThread.Start()
            Log($"Human RapidFire gestartet - Pattern: {CurrentPattern}, Stufe: {RapidFireSpeedFactor}")
        End Sub

        Public Sub StopRapidFire()
            RapidFireStopEvent.Set()

            If RapidFireThread IsNot Nothing AndAlso RapidFireThread.IsAlive Then
                If Not RapidFireThread.Join(500) Then
                    RapidFireThread.Abort()
                End If
            End If

            ' Stelle sicher, dass die linke Maustaste released wird
            If LeftButtonPressedFromRapidFire Then
                Release(MakcuMouseButton.Left)
                LeftButtonPressedFromRapidFire = False
                Log("RapidFire: Maustaste released")
            End If

            RapidFireThread = Nothing
            Log("Human RapidFire gestoppt")
        End Sub

        Public Sub RapidFireWorker()
            Dim stopwatch As New Stopwatch()
            stopwatch.Start()

            While Not RapidFireStopEvent.IsSet
                Try
                    ' Prüfe ob linke Maustaste gedrückt ist (vom Benutzer)
                    Dim leftButtonPressed As Boolean = False
                    SyncLock _buttonStates
                        leftButtonPressed = _buttonStates(MakcuMouseButton.Left)
                    End SyncLock

                    If leftButtonPressed AndAlso Not LeftButtonPressedFromRapidFire Then
                        ' Erster Klick - drücke die Taste
                        Press(MakcuMouseButton.Left)
                        LeftButtonPressedFromRapidFire = True
                        LastClickTime = stopwatch.ElapsedMilliseconds

                    ElseIf leftButtonPressed AndAlso LeftButtonPressedFromRapidFire Then
                        ' RapidFire aktiv - menschliche Klicks
                        Dim currentTime = stopwatch.ElapsedMilliseconds
                        Dim timeSinceLastClick = currentTime - LastClickTime
                        Dim nextInterval = GetHumanizedInterval()

                        If timeSinceLastClick >= nextInterval Then
                            ' Human-like Klick: Release und Press mit variabler Dauer
                            Release(MakcuMouseButton.Left)

                            ' Kurze Pause zwischen Release und nächstem Press
                            Dim clickDuration = GetHumanizedClickDuration()
                            Thread.Sleep(Math.Max(1, clickDuration))

                            Press(MakcuMouseButton.Left)
                            LastClickTime = stopwatch.ElapsedMilliseconds
                        End If

                    ElseIf Not leftButtonPressed AndAlso LeftButtonPressedFromRapidFire Then
                        ' Benutzer hat Taste losgelassen - release final
                        Release(MakcuMouseButton.Left)
                        LeftButtonPressedFromRapidFire = False
                        Exit While ' Sofort beenden wenn Taste losgelassen
                    End If

                    ' Kurze Pause um CPU zu schonen
                    Thread.Sleep(1)

                Catch ex As ThreadAbortException
                    Exit While
                Catch ex As Exception
                    Log($"Fehler in RapidFire Worker: {ex.Message}")
                    Thread.Sleep(10)
                End Try
            End While

            ' Cleanup beim Beenden
            If LeftButtonPressedFromRapidFire Then
                Release(MakcuMouseButton.Left)
                LeftButtonPressedFromRapidFire = False
            End If

            stopwatch.Stop()
        End Sub

        ' Gaming Profile Methoden - PUBLIC
        Public Sub SetGamingProfile(profile As Integer)
            Select Case profile
                Case 1 ' Langsam
                    CurrentPattern = "slow"
                    RapidFireSpeedFactor = 1
                Case 2 ' Mittel
                    CurrentPattern = "medium"
                    RapidFireSpeedFactor = 2
                Case 3 ' Schnell
                    CurrentPattern = "fast"
                    RapidFireSpeedFactor = 3
                Case Else
                    CurrentPattern = "medium"
                    RapidFireSpeedFactor = 2
            End Select
            Log($"Gaming Profile Stufe {profile} aktiviert - Pattern: {CurrentPattern}")
        End Sub

        Public Sub SetBurstMode()
            CurrentPattern = "burst"
            RapidFireSpeedFactor = 3
            Log("Burst Mode aktiviert")
        End Sub

        ' Toggle RapidFire - PUBLIC
        Public Sub ToggleRapidFire()
            RapidFireEnabled = Not RapidFireEnabled
            If RapidFireEnabled Then
                StartRapidFire()
            Else
                StopRapidFire()
            End If
            Log($"RapidFire: {If(RapidFireEnabled, "AKTIV", "INAKTIV")}")
        End Sub

        ' Automatische Steuerung bei Tastendruck - PUBLIC
        Public Sub HandleRapidFireButtonPress(button As MakcuMouseButton, isPressed As Boolean)
            If button = MakcuMouseButton.Left Then
                If isPressed And RapidFireEnabled Then
                    ' RapidFire startet automatisch durch den Worker
                ElseIf Not isPressed Then
                    ' Sofort stoppen wenn Taste losgelassen
                    If LeftButtonPressedFromRapidFire Then
                        Release(MakcuMouseButton.Left)
                        LeftButtonPressedFromRapidFire = False
                    End If
                End If
            End If
        End Sub

    End Class

End Namespace
