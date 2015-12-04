Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Text
Imports ViewModels.Base

Namespace DataAccessLayer

    Public Class TextFileAccess
        Inherits UINotificationBase

#Region "Properties"

        Private _currentFileName As String
        ''' <summary>
        ''' The Name Of The Currently Opened Text File Or the File To Open.
        ''' </summary>
        Public Property CurrentFileName() As String
            Get
                Return _currentFileName
            End Get
            Set(ByVal value As String)
                _currentFileName = value
                ReportPropertyChanged("CurrentFileName")
            End Set
        End Property

        Private _FileContent As List(Of String) = New List(Of String)
        ''' <summary>
        ''' Strongly Typed Object That Contains Each Line Of The Text File.
        ''' </summary>
        Public ReadOnly Property FileContent() As List(Of String)
            Get
                Return _FileContent
            End Get
        End Property

        Private _DialogTitle As String = "Open File"
        ''' <summary>
        ''' The Title To Be Displayed On The OpenFileDialog Window
        ''' </summary>
        Public Property DialogTitle() As String
            Get
                Return _DialogTitle
            End Get
            Set(ByVal value As String)
                _DialogTitle = value
            End Set
        End Property

        Private _DialogFilter As String = "All Files|*.*"
        ''' <summary>
        ''' The Options For Which File Types To Display In The OpenFileDialog Window
        ''' </summary>
        Public Property DialogFilter() As String
            Get
                Return _DialogFilter
            End Get
            Set(ByVal value As String)
                _DialogFilter = value
            End Set
        End Property

        Private _DefaultExtension As String
        ''' <summary>
        ''' The Default File Extension
        ''' </summary>
        Public Property DefaultExtension() As String
            Get
                Return _DefaultExtension
            End Get
            Set(ByVal value As String)
                _DefaultExtension = value
            End Set
        End Property

#End Region

#Region "Constructors"

        ''' <summary>
        ''' Initialises The Class But Does Not Read The Text File.
        ''' </summary>
        ''' <param name="pDialogTitle">The Title To Be Displayed On The OpenFileDialog Window</param>
        ''' <param name="pDefaultExt">The Default File Extension</param>
        ''' <param name="pDialogFilter">The Options For Which File Types To Display In The OpenFileDialog Window</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pDialogTitle As String, ByVal pDefaultExt As String, ByVal pDialogFilter As String)
            _DialogTitle = pDialogTitle
            _DialogFilter = pDialogFilter
        End Sub

#End Region

#Region "Validation"

        ''' <summary>
        ''' Indicates If The File Actually Exists.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FileExists() As Boolean
            Return IO.File.Exists(_currentFileName)
        End Function

        ''' <summary>
        ''' Checks That We Are Not Going To Try To OverWrite Lines That Are There.
        ''' </summary>
        ''' <param name="range">The High End Count Value To Compare.</param>
        ''' <returns>True If Ok, Fasle Otherwise.</returns>
        Public Function WithinBounds(ByVal range As Integer) As Boolean
            Return Me.FileContent.Count >= range
        End Function

#End Region

#Region "Methods"

        ''' <summary>
        ''' Creates A New Text File.
        ''' </summary>
        Public Function CreateFile() As Boolean

            Try
                ' Variables
                Dim result As Nullable(Of Boolean) = Nothing
                Dim tempName As String = String.Empty
                Dim sfd As Microsoft.Win32.SaveFileDialog = New Microsoft.Win32.SaveFileDialog

                ' Initialise Save File Dialog And Open It.
                sfd.Title = "New " & _DialogTitle
                sfd.DefaultExt = _DefaultExtension
                sfd.Filter = _DialogFilter
                result = sfd.ShowDialog

                ' If User Saved
                If result = True Then
                    ' Assign Name Tempoarily
                    tempName = sfd.FileName
                    ' Create New FileInfo Object
                    Dim fi As New FileInfo(tempName)
                    Try
                        ' Create The File.
                        fi.CreateText()
                        'Assign Name Permanently.
                        Me.CurrentFileName = tempName
                        ' Clear All Objects Form Our Contents List.
                        Me.FileContent.Clear()
                        Return True
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString, "Error Creating New File", MessageBoxButton.OK, MessageBoxImage.Error)
                        Return False
                    End Try
                End If
                Return True
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error Saving New File", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Calls The OpenFileDialog And Initiates Reading.
        ''' </summary>
        ''' <returns>True If The Operation Was Successful or False If Not.</returns>
        Public Function OpenFile() As Boolean
            Try
                ' Get The File Name
                Dim ofd As Microsoft.Win32.OpenFileDialog = New Microsoft.Win32.OpenFileDialog()
                ofd.Title = "Open " & _DialogTitle
                ofd.Filter = _DialogFilter
                ofd.ShowDialog()
                If Not String.IsNullOrEmpty(ofd.FileName) Then
                    Me.CurrentFileName = ofd.FileName
                End If

                'Read The File
                Me.Read()

                Return True
            Catch ex As Exception
                System.Windows.MessageBox.Show(ex.ToString, "Error Opening File", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Permanently Deletes A File.
        ''' </summary>
        Public Function DeleteFile() As Boolean
            Try
                If Me.FileExists Then
                    Dim fi As New FileInfo(_currentFileName)
                    fi.Delete()
                    Me.FileContent.Clear()
                    Me.CurrentFileName = ""
                End If
                Return True
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error Deleting File", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Reads All Lines From A Text File Into A Strongly Typed List(Of String)
        ''' </summary>
        ''' <returns>True If Successful, False Otherwise</returns>
        ''' <remarks></remarks>
        Private Function Read() As Boolean
            If Me.FileExists Then
                Try
                    ' Clear All Items In Out List First.
                    Me.FileContent.Clear()

                    ' Read The File, Adding Wach Line To Our FileContent List
                    Using sr As StreamReader = IO.File.OpenText(_currentFileName)
                        Do While sr.Peek <> -1
                            _FileContent.Add(sr.ReadLine)
                        Loop
                    End Using

                    Return True
                Catch ex As Exception
                    System.Windows.MessageBox.Show(ex.ToString, "Error Reading Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Appends A Line or Lines To The End Of The Text File.
        ''' </summary>
        Public Function Append(ByVal ParamArray lines() As String) As Boolean
            If Me.FileExists Then
                Try
                    ' Add Lines To Both The List And File.
                    Using sw As StreamWriter = New StreamWriter(_currentFileName, True)
                        For Each Line As String In lines
                            _FileContent.Add(Line)
                            sw.WriteLine(Line)
                        Next
                    End Using

                    Return True
                Catch ex As Exception
                    System.Windows.MessageBox.Show(ex.ToString, "Error Appending To Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Changes The Line Or Lines In A Text File For A Given 'Zero Based' Ordinal Position.
        ''' </summary>
        ''' <param name="index">Integer Representing The 'Zero Based' Ordinal Position From Which To Start Writing.</param>
        ''' <param name="Lines">String Array Containing The Lines In Which To Write To The File</param>
        ''' <returns>True If Operation Was Successful, False Otherwise.</returns>
        ''' <remarks></remarks>
        Public Function Change(ByVal index As Integer, ByVal ParamArray Lines() As String) As Boolean

            ' Firstly - Check That We Are Not Going To Try To OverWrite Lines That Are Not There.
            If Me.FileExists And Me.WithinBounds(index + Lines.Count) Then
                Try

                    ' Secondly - Over Write Items In The List.
                    Dim lineCounter As Integer = 0
                    For indexCounter = index To (index + Lines.Count) - 1
                        Me.FileContent.Item(indexCounter) = Lines(lineCounter)
                        lineCounter += 1
                    Next

                    ' Thirdly - Write The Entire Contents Of The List To File.
                    Me.Write()

                    Return True
                Catch ex As Exception
                    System.Windows.MessageBox.Show(ex.ToString, "Error Making Changes To Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            Else
                System.Windows.MessageBox.Show("Either The File Does Not Exist Or You Are Trying To Change Too Many Lines.", "Error Making Changes To Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If

        End Function

        ''' <summary>
        ''' Removes A Single Specific Line From The Text File According To A 'Zero Based' Index Ordinal Arguement.
        ''' </summary>
        ''' <param name="Index">The 'Zero Based' Ordinal Position Of The Line To Remove.</param>
        ''' <returns>True If Successful, Otherwise False.</returns>
        ''' <remarks></remarks>
        Public Function Remove(ByVal Index As Integer)
            If Me.FileExists Then
                Try
                    ' Remove The Item/Line.
                    Me.FileContent.RemoveAt(Index)

                    ' Write The Entire Contents Of The List To File.
                    Me.Write()

                    Return True
                Catch ex As Exception
                    System.Windows.MessageBox.Show(ex.ToString, "Error Removing Line From Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            Else
                System.Windows.MessageBox.Show("The File Does Not Exist.", "Error Removing Lines From Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If
        End Function

        ''' <summary>
        ''' Removes A Range Of Lines From The Text File According To A 'Zero Based' Index Ordinal Arguement.
        ''' </summary>
        ''' <param name="startIndex">The 'Zero Based' Ordinal Position Of Where To Start Removing Lines From.</param>
        ''' <param name="count">The Number Of Lines To Remove Including Index Position.</param>
        ''' <returns>True If Successful, Otherwise False.</returns>
        ''' <remarks>The Line Corresponding To The 'startIndex' Is Also Removed.</remarks>
        Public Function Remove(ByVal startIndex As Integer, ByVal count As Integer)

            ' Firstly - Check That We Are Not Going To Try To OverWrite Lines That Are Not There.
            If Me.FileExists And Me.WithinBounds(startIndex + count) Then '(Me.FileContent.Count - 1) >= count Then
                Try

                    ' Secondly - Remove The Range Of Items / Lines.
                    Me.FileContent.RemoveRange(startIndex, count)

                    ' Thirdly - Write The Entire Contents Of The List To File.
                    Me.Write()

                    Return True
                Catch ex As Exception
                    System.Windows.MessageBox.Show(ex.ToString, "Error Removing Lines From Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            Else
                System.Windows.MessageBox.Show("Either The File Does Not Exist Or You Are Trying To Remove Too Many Lines.", "Error Removing Lines From Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If
        End Function

        ''' <summary>
        ''' Writes All Contents To The Text File.
        ''' </summary>
        ''' <returns>True If Successful, Otherwise False.</returns>
        Private Function Write() As Boolean
            If Me.FileExists Then
                Try
                    Using sw As StreamWriter = New StreamWriter(_currentFileName)
                        For Each Line As String In Me.FileContent
                            sw.WriteLine(Line)
                        Next
                    End Using
                    Return True
                Catch ex As Exception
                    System.Windows.MessageBox.Show(ex.ToString, "Error Writing To Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            Else
                System.Windows.MessageBox.Show("The File Does Not Exist.", "Error Removing Lines From Text File", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If
        End Function

#End Region

    End Class

End Namespace
