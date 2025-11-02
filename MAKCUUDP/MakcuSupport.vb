Public Class MakcuSupport
    Private Shared _makcuInstance As MouseMovementLibraries.MakcuSupport.MakcuMouse = Nothing

    Public Shared ReadOnly Property MakcuInstance() As MouseMovementLibraries.MakcuSupport.MakcuMouse
        Get
            Return _makcuInstance
        End Get
    End Property

    Private Shared _isMakcuLoaded As System.Boolean = False
    Private Shared _isSubscribedToButtonEvents As System.Boolean = False

    Private Const DefaultDebugLoggingForInternalCreation As System.Boolean = False
    Private Const DefaultSendInitCommandsForInternalCreation As System.Boolean = True

    Public Shared Sub ConfigureMakcuInstance(ByVal debugEnabled As System.Boolean, ByVal sendInitCmds As System.Boolean)
        System.Console.WriteLine("MakcuMain: Configuring MakcuInstance. Debug: " & debugEnabled.ToString() &
                                     ", SendInitCmds: " & sendInitCmds.ToString())

        UnsubscribeFromButtonEvents()

        If _makcuInstance IsNot Nothing Then
            _makcuInstance.Dispose()
        End If

        _makcuInstance = New MouseMovementLibraries.MakcuSupport.MakcuMouse(debugEnabled, sendInitCmds)
        _isMakcuLoaded = False
    End Sub

    Private Shared Async Function InitializeMakcuDevice() As System.Threading.Tasks.Task(Of System.Boolean)
        If _isMakcuLoaded AndAlso _makcuInstance IsNot Nothing AndAlso _makcuInstance.IsInitializedAndConnected Then
            System.Console.WriteLine("MakcuMain: InitializeMakcuDevice called, but Makcu is already loaded and connected.")
            Console.WriteLine("MAKCU instance initialized for Port=" & _makcuInstance.PortName, 5000)
            Return True
        End If

        If _makcuInstance Is Nothing Then
            ConfigureMakcuInstance(DefaultDebugLoggingForInternalCreation, DefaultSendInitCommandsForInternalCreation)
        End If

        Try
            If _makcuInstance Is Nothing OrElse Not _makcuInstance.Init() Then
                Console.WriteLine("MAKCU initialization failed." & Environment.NewLine &
                                                   "Verify that the device is connected and not in use by another application.")
                _isMakcuLoaded = False
                Return False
            End If

            Dim version As System.String = _makcuInstance.GetKmVersion()

            If System.String.IsNullOrWhiteSpace(version) Then
                Console.WriteLine("No version response received from the Makcu device on " & _makcuInstance.PortName & "." & Environment.NewLine &
                                                   "Ensure the firmware is compatible and responds to the 'km.version()' command." & Environment.NewLine &
                                                   "The connection might be unstable or the device is not responding as expected.")
            End If

            _isMakcuLoaded = True

            SubscribeToButtonEvents()


            Console.WriteLine("MAKCU instance initialized for Port=" & _makcuInstance.PortName, 5000)

            Return True

        Catch ex As System.Exception
            Console.WriteLine("Catastrophic exception during Makcu initialization. Error: " & ex.Message &
                                               Environment.NewLine & "Stack Trace: " & ex.StackTrace)
            _isMakcuLoaded = False

            If _makcuInstance IsNot Nothing AndAlso _makcuInstance.IsInitializedAndConnected Then
                _makcuInstance.Close()
            End If

            Return False
        End Try
    End Function

    Public Shared Async Function Load() As System.Threading.Tasks.Task(Of System.Boolean)
        Return Await InitializeMakcuDevice().ConfigureAwait(False)
    End Function

    Public Shared Sub Unload()
        UnsubscribeFromButtonEvents()

        If _makcuInstance IsNot Nothing Then
            _makcuInstance.Close()
        End If

        _isMakcuLoaded = False
        System.Console.WriteLine("MakcuMain: Makcu device unloaded/closed.")
    End Sub

    Public Shared Sub DisposeInstance()
        Unload()

        If _makcuInstance IsNot Nothing Then
            _makcuInstance.Dispose()
        End If

        _makcuInstance = Nothing
        _isMakcuLoaded = False
        System.Console.WriteLine("MakcuMain: MakcuMouse instance disposed (null).")
    End Sub

    Private Shared Sub SubscribeToButtonEvents()
        If _makcuInstance IsNot Nothing AndAlso Not _isSubscribedToButtonEvents Then
            AddHandler _makcuInstance.ButtonStateChanged, AddressOf OnMakcuButtonStateChanged
            _isSubscribedToButtonEvents = True
            System.Diagnostics.Debug.WriteLine("MakcuMain: Subscribed to MakcuInstance.ButtonStateChanged events.")
        End If
    End Sub

    Private Shared Sub UnsubscribeFromButtonEvents()
        If _makcuInstance IsNot Nothing AndAlso _isSubscribedToButtonEvents Then
            RemoveHandler _makcuInstance.ButtonStateChanged, AddressOf OnMakcuButtonStateChanged
            _isSubscribedToButtonEvents = False
            System.Diagnostics.Debug.WriteLine("MakcuMain: Unsubscribed from MakcuInstance.ButtonStateChanged events.")
        End If
    End Sub

    Private Shared Sub OnMakcuButtonStateChanged(ByVal button As MouseMovementLibraries.MakcuSupport.MakcuMouseButton, ByVal isPressed As System.Boolean)
        Dim state As System.String = If(isPressed, "Pressed", "Released")
        System.Diagnostics.Debug.WriteLine(button.ToString() & " physical " & state & "!")
    End Sub

    Private Sub New()
    End Sub

End Class
