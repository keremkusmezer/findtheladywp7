Partial Public Class App
    Inherits Application

    Public Property RootFrame As PhoneApplicationFrame

    Public Sub New()
        InitializeComponent()

        InitializePhoneApplication()

        If System.Diagnostics.Debugger.IsAttached Then
            Application.Current.Host.Settings.EnableFrameRateCounter = True
            PhoneApplicationService.Current.UserIdleDetectionMode = IdleDetectionMode.Disabled
        End If

    End Sub

    Private Sub Application_Launching(ByVal sender As Object, ByVal e As LaunchingEventArgs)

    End Sub

    Private Sub Application_Activated(ByVal sender As Object, ByVal e As ActivatedEventArgs)

    End Sub

    Private Sub Application_Deactivated(ByVal sender As Object, ByVal e As DeactivatedEventArgs)

    End Sub

    Private Sub Application_Closing(ByVal sender As Object, ByVal e As ClosingEventArgs)

    End Sub

    Private Sub RootFrame_NavigationFailed(ByVal sender As Object, ByVal e As NavigationFailedEventArgs)
        If Diagnostics.Debugger.IsAttached Then
            Diagnostics.Debugger.Break()
        End If
    End Sub

    Public Sub Application_UnhandledException(ByVal sender As Object, ByVal e As ApplicationUnhandledExceptionEventArgs) Handles Me.UnhandledException

        If Diagnostics.Debugger.IsAttached Then
            Diagnostics.Debugger.Break()
        Else
            e.Handled = True
            MessageBox.Show(e.ExceptionObject.Message & Environment.NewLine & e.ExceptionObject.StackTrace,
                            "Error", MessageBoxButton.OK)
        End If
    End Sub

#Region "Phone application initialization"

    Private phoneApplicationInitialized As Boolean = False

    Private Sub InitializePhoneApplication()
        If phoneApplicationInitialized Then
            Return
        End If

        RootFrame = New PhoneApplicationFrame()
        AddHandler RootFrame.Navigated, AddressOf CompleteInitializePhoneApplication

        AddHandler RootFrame.NavigationFailed, AddressOf RootFrame_NavigationFailed

        phoneApplicationInitialized = True
    End Sub

    Private Sub CompleteInitializePhoneApplication(ByVal sender As Object, ByVal e As NavigationEventArgs)

        If RootVisual IsNot RootFrame Then
            RootVisual = RootFrame
        End If

        RemoveHandler RootFrame.Navigated, AddressOf CompleteInitializePhoneApplication
    End Sub
#End Region

End Class