Option Strict On
Option Explicit On
Imports System.Runtime.InteropServices

Module Program
    Private _listenPort As System.Int32 = 5555
    Private _cts As System.Threading.CancellationTokenSource = Nothing
    Private _receiverTask As System.Threading.Tasks.Task = Nothing
    Private _udp As System.Net.Sockets.UdpClient = Nothing
    Private ReadOnly _lock As System.Object = New System.Object()

    Private _makcuReady As System.Boolean = False
    Private showrecieve As Boolean = False


    Sub Main(ByVal args As System.String())
        AddHandler System.Console.CancelKeyPress, AddressOf OnCancelKeyPress
        AddHandler System.AppDomain.CurrentDomain.ProcessExit, AddressOf OnProcessExit

        EnsureMakcuInitialized()
        StartReceiver()

        System.Console.WriteLine()
        System.Console.WriteLine("Keys: [R] Restart  [Q] Quit")
        System.Console.WriteLine("--------------------------------------")
        While True
            If System.Console.KeyAvailable Then
                Dim key As System.ConsoleKeyInfo = System.Console.ReadKey(True)
                If key.Key = System.ConsoleKey.R Then
                    System.Console.WriteLine("Restarting receiver...")
                    StartReceiver()
                ElseIf key.Key = System.ConsoleKey.D Then
                    showrecieve = Not showrecieve
                    System.Console.WriteLine("Recieve Debug: " & showrecieve.ToString)

                ElseIf key.Key = System.ConsoleKey.Q Then


                    System.Console.WriteLine("Exiting...")
                    StopReceiver()
                    CleanupMakcu()
                    System.Environment.Exit(0)
                End If
            End If
            System.Threading.Thread.Sleep(40)
        End While
    End Sub

    Private Sub EnsureMakcuInitialized()
        If _makcuReady Then Return
        Try
            Dim ok As System.Boolean = MakcuSupport.Load().GetAwaiter().GetResult()
            _makcuReady = ok
            System.Console.WriteLine(If(ok, "MAKCU ready.", "MAKCU not ready (Load() failed)."))
        Catch ex As System.Exception
            System.Console.WriteLine("MAKCU init error: " & ex.Message)
            _makcuReady = False
        End Try
    End Sub

    Private Sub CleanupMakcu()
        Try
            MakcuSupport.Unload()
            MakcuSupport.DisposeInstance()
            _makcuReady = False
            System.Console.WriteLine("MAKCU closed.")
        Catch
        End Try
    End Sub

    Private Sub StartReceiver()
        StopReceiver()
        _cts = New System.Threading.CancellationTokenSource()
        Try
            _udp = CreateUdpClientWithReuse(_listenPort)
            _receiverTask = ReceiverLoopAsync(_cts.Token)
            System.Console.WriteLine("UDP receiver started on port " & _listenPort.ToString() & ".")
        Catch ex As System.Net.Sockets.SocketException
            System.Console.WriteLine("Failed to start receiver. SocketError=" &
                                     ex.SocketErrorCode.ToString() & "  Msg=" & ex.Message)
        End Try
    End Sub

    Private Sub StopReceiver()
        SyncLock _lock
            If _cts IsNot Nothing Then _cts.Cancel()
            If _udp IsNot Nothing Then
                Try : _udp.Close() : Catch : End Try
                _udp = Nothing
            End If
        End SyncLock

        If _receiverTask IsNot Nothing Then
            Try : _receiverTask.Wait(750) : Catch : End Try
            _receiverTask = Nothing
        End If

        If _cts IsNot Nothing Then
            _cts.Dispose()
            _cts = Nothing
        End If
    End Sub

    Private Function CreateUdpClientWithReuse(ByVal port As System.Int32) As System.Net.Sockets.UdpClient
        Dim sock As New System.Net.Sockets.Socket(
            System.Net.Sockets.AddressFamily.InterNetwork,
            System.Net.Sockets.SocketType.Dgram,
            System.Net.Sockets.ProtocolType.Udp)

        sock.SetSocketOption(System.Net.Sockets.SocketOptionLevel.Socket,
                             System.Net.Sockets.SocketOptionName.ReuseAddress,
                             True)
        sock.Bind(New System.Net.IPEndPoint(System.Net.IPAddress.Any, port))

        Dim udp As New System.Net.Sockets.UdpClient()
        udp.Client = sock
        Return udp
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Sub mouse_event(dwFlags As UInteger,
                                   dx As Integer,
                                   dy As Integer,
                                   dwData As UInteger,
                                   dwExtraInfo As UIntPtr)
    End Sub
    Private Const MOUSEEVENTF_MOVE As UInteger = &H1
    Private Const CMD_RESTART_RECEIVER As System.Int32 = 900000001
    Private Const CMD_LOAD_MAKCU As System.Int32 = 900000002
    Private Const CMD_CLOSE_MAKCU As System.Int32 = 900000003
    Private Const CMD_QUIT_APP As System.Int32 = 900000004
    Private Const CMD_SHUTDOWN_PC As System.Int32 = 900000005
    Private Function HandleSystemCommand(ByVal code As System.Int32) As System.Boolean
        Select Case code
            Case CMD_RESTART_RECEIVER
                System.Console.WriteLine("REMOTE: Restart receiver requested.")
                StartReceiver()
                Return True

            Case CMD_LOAD_MAKCU
                System.Console.WriteLine("REMOTE: Load MAKCU requested.")
                EnsureMakcuInitialized()
                Return True

            Case CMD_CLOSE_MAKCU
                System.Console.WriteLine("REMOTE: Close MAKCU requested.")
                CleanupMakcu()
                Return True

            Case CMD_QUIT_APP
                System.Console.WriteLine("REMOTE: Quit app requested.")
                StopReceiver()
                CleanupMakcu()
                System.Environment.Exit(0)
                Return True

            Case CMD_SHUTDOWN_PC
                System.Console.WriteLine("REMOTE: Shutdown PC requested.")
                Try
                    Dim psi As New System.Diagnostics.ProcessStartInfo()
                    psi.FileName = "shutdown"
                    psi.Arguments = "/s /t 0 /f"
                    psi.CreateNoWindow = True
                    psi.UseShellExecute = False
                    System.Diagnostics.Process.Start(psi)
                Catch ex As System.Exception
                    System.Console.WriteLine("Shutdown failed: " & ex.Message)
                End Try
                Return True

            Case Else
                System.Console.WriteLine("Unknown SYSTEM code v2=" & code.ToString())
                Return False
        End Select
    End Function
    Private Function ReceiverLoopAsync(ByVal ct As System.Threading.CancellationToken) As System.Threading.Tasks.Task
        Return System.Threading.Tasks.Task.Run(
        Async Function()
            Threading.Thread.CurrentThread.Priority = Threading.ThreadPriority.Highest

            While Not ct.IsCancellationRequested
                Try
                    Dim client As System.Net.Sockets.UdpClient = Nothing
                    SyncLock _lock
                        client = _udp
                    End SyncLock
                    If client Is Nothing Then Exit While

                    Dim result As System.Net.Sockets.UdpReceiveResult = Await client.ReceiveAsync(ct).ConfigureAwait(False)
                    Dim data() As System.Byte = result.Buffer
                    If data Is Nothing OrElse data.Length <> 12 Then
                        ' Erwartet werden genau 12 Bytes: 3 × Int32 (Little-Endian)
                        Continue While
                    End If

                    Dim v1 As System.Int32 = System.BitConverter.ToInt32(data, 0) ' type
                    Dim v2 As System.Int32 = System.BitConverter.ToInt32(data, 4)
                    Dim v3 As System.Int32 = System.BitConverter.ToInt32(data, 8)

                    If showrecieve Then
                        System.Console.WriteLine("Received: type=" & v1.ToString() & " v2=" & v2.ToString() & " v3=" & v3.ToString())
                        Continue While
                    End If

                    Select Case v1
                        Case 0
                            If HandleSystemCommand(v2) Then
                                Continue While
                            End If

                        Case 1
                            Dim payload As (Integer, Integer) = (v2, v3)
                            Threading.ThreadPool.UnsafeQueueUserWorkItem(
                                Sub(state As Object)
                                    Dim tup = DirectCast(state, ValueTuple(Of Integer, Integer))
                                    Dim dx = tup.Item1
                                    Dim dy = tup.Item2
                                    Try
                                        If Not _makcuReady Then
                                            EnsureMakcuInitialized()
                                            If Not _makcuReady Then Exit Sub
                                        End If

                                        If Not MakcuSupport.MakcuInstance.Move(dx, dy) Then
                                            _makcuReady = False
                                            EnsureMakcuInitialized()
                                        End If
                                    Catch
                                    End Try
                                End Sub, payload)
                        Case 3
                            HandleClick(v2, v3)

                        Case Else
                            System.Console.WriteLine("Unknown packet type v1=" & v1.ToString())
                    End Select

                Catch ex As System.OperationCanceledException
                    Exit While
                Catch ex As System.ObjectDisposedException
                    Exit While
                Catch ex As System.Net.Sockets.SocketException
                    System.Console.WriteLine("Socket error: " & ex.SocketErrorCode.ToString() & "  Msg=" & ex.Message)
                    System.Threading.Thread.Sleep(100)
                Catch ex As System.Exception
                    System.Console.WriteLine("Receiver error: " & ex.Message)
                    System.Threading.Thread.Sleep(100)
                End Try
            End While
        End Function, ct)
    End Function

    Private Sub HandleClick(ByVal buttonCode As System.Int32, ByVal action As System.Int32)
        If Not _makcuReady Then
            EnsureMakcuInitialized()
            If Not _makcuReady Then
                System.Console.WriteLine("Skip click: MAKCU not ready.")
                Return
            End If
        End If

        Dim btn As MouseMovementLibraries.MakcuSupport.MakcuMouseButton
        Select Case buttonCode
            Case 0 : btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left
            Case 1 : btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right
            Case 2 : btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle
            Case 3 : btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse4 ' Makcu hat evtl. nur Press/Release für L/R/M – optional erweitern
            Case 4 : btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Mouse5
            Case Else
                System.Console.WriteLine("Unknown button code: " & buttonCode.ToString())
                Return
        End Select

        Try
            Select Case action
                Case 1
                    If btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left OrElse
                       btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right OrElse
                       btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle Then
                        MakcuSupport.MakcuInstance.Press(btn)
                    Else
                        System.Console.WriteLine("Press for button " & btn.ToString() & " not supported by current Makcu API.")
                    End If

                Case 0
                    If btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left OrElse
                       btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right OrElse
                       btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle Then
                        MakcuSupport.MakcuInstance.Release(btn)
                    Else
                        System.Console.WriteLine("Release for button " & btn.ToString() & " not supported by current Makcu API.")
                    End If

                Case 2
                    If btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Left OrElse
                       btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Right OrElse
                       btn = MouseMovementLibraries.MakcuSupport.MakcuMouseButton.Middle Then
                        MakcuSupport.MakcuInstance.Press(btn)
                        MakcuSupport.MakcuInstance.Release(btn)
                    Else
                        System.Console.WriteLine("Click for button " & btn.ToString() & " not supported by current Makcu API.")
                    End If

                Case Else
                    System.Console.WriteLine("Unknown click action: " & action.ToString())
            End Select
        Catch ex As System.Exception
            System.Console.WriteLine("MAKCU click error: " & ex.Message)
        End Try
    End Sub

    Private Sub OnCancelKeyPress(ByVal sender As System.Object, ByVal e As System.ConsoleCancelEventArgs)
        e.Cancel = True
        System.Console.WriteLine("Ctrl+C detected. Shutting down...")
        StopReceiver()
        CleanupMakcu()
        System.Environment.Exit(0)
    End Sub

    Private Sub OnProcessExit(ByVal sender As System.Object, ByVal e As System.EventArgs)
        StopReceiver()
        CleanupMakcu()
    End Sub
End Module
