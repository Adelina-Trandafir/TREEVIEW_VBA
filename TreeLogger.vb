Imports System.IO

''' <summary>
''' Logger centralizat bazat pe treeId. 
''' Fișierul se suprascrie la fiecare pornire.
''' Thread-safe prin SyncLock.
''' </summary>
Public Class TreeLogger
    Private Shared _logPath As String = Nothing
    Private Shared ReadOnly _lock As New Object()
    Private Shared _initialized As Boolean = False
    Private Shared _startTime As DateTime

    Public Enum LogLevel
        DEBUG_ = 0
        INFO = 1
        WARN = 2
        ERR = 3
    End Enum

    ''' <summary>
    ''' Inițializează logger-ul. Apelat O SINGURĂ DATĂ la pornirea aplicației.
    ''' Suprascrie fișierul existent.
    ''' </summary>
    Public Shared Sub Init(treeId As String)
        SyncLock _lock
            If _initialized Then Return

            _startTime = DateTime.Now

            ' Construim calea: folder exe + log_{treeId}.txt
            Dim folder As String = AppDomain.CurrentDomain.BaseDirectory
            Dim safeName As String = If(String.IsNullOrEmpty(treeId), "unknown", SanitizeFileName(treeId))
            _logPath = Path.Combine(folder, $"log_{safeName}.txt")

            Try
                ' Suprascrie fișierul (creează nou la fiecare pornire)
                File.WriteAllText(_logPath,
                    $"========================================{Environment.NewLine}" &
                    $"  TREEVIEW_VBA Log - {treeId}{Environment.NewLine}" &
                    $"  Start: {_startTime:yyyy-MM-dd HH:mm:ss.fff}{Environment.NewLine}" &
                    $"  Machine: {Environment.MachineName}{Environment.NewLine}" &
                    $"========================================{Environment.NewLine}")

                _initialized = True
            Catch ex As Exception
                ' Fallback: dacă nu putem scrie în folderul exe, încercăm Temp
                Try
                    folder = Path.GetTempPath()
                    _logPath = Path.Combine(folder, $"log_{safeName}.txt")
                    File.WriteAllText(_logPath, $"[FALLBACK] Log start: {_startTime:yyyy-MM-dd HH:mm:ss.fff}{Environment.NewLine}")
                    _initialized = True
                Catch
                    ' Nu putem loga nicăieri - continuăm silențios
                    _initialized = False
                End Try
            End Try
        End SyncLock
    End Sub

    ' ─── Metode publice de logare ───

    Public Shared Sub Debug(message As String, Optional source As String = "", Optional dummy1 As Object = Nothing, Optional dummy2 As Object = Nothing)
        Write(LogLevel.DEBUG_, message, source)
    End Sub

    Public Shared Sub Info(message As String, Optional source As String = "", Optional dummy1 As Object = Nothing, Optional dummy2 As Object = Nothing)
        Write(LogLevel.INFO, message, source)
    End Sub

    Public Shared Sub Warn(message As String, Optional source As String = "", Optional dummy1 As Object = Nothing, Optional dummy2 As Object = Nothing)
        Write(LogLevel.WARN, message, source)
    End Sub

    Public Shared Sub Err(message As String, Optional source As String = "", Optional dummy1 As Object = Nothing, Optional dummy2 As Object = Nothing)
        Write(LogLevel.ERR, message, source)
    End Sub

    ''' <summary>
    ''' Loghează o excepție cu stack trace.
    ''' </summary>
    Public Shared Sub Ex(ex As Exception, Optional source As String = "", Optional dummy1 As Object = Nothing, Optional dummy2 As Object = Nothing)
        If ex Is Nothing Then Return
        Dim msg As String = $"{ex.GetType().Name}: {ex.Message}{Environment.NewLine}  StackTrace: {ex.StackTrace}"
        If ex.InnerException IsNot Nothing Then
            msg &= $"{Environment.NewLine}  Inner: {ex.InnerException.Message}"
        End If
        Write(LogLevel.ERR, msg, source)
    End Sub

    ''' <summary>
    ''' Loghează durata unei operații (pentru profiling).
    ''' </summary>
    Public Shared Sub Perf(operation As String, elapsedMs As Long, Optional source As String = "")
        Write(LogLevel.DEBUG_, $"PERF [{operation}] {elapsedMs}ms", source)
    End Sub

    ' ─── Implementare internă ───

    Private Shared Sub Write(level As LogLevel, message As String, source As String)
        If Not _initialized Then Return

        Dim elapsed As TimeSpan = DateTime.Now - _startTime
        Dim levelStr As String = level.ToString().TrimEnd("_"c).PadRight(5)
        Dim srcStr As String = If(String.IsNullOrEmpty(source), "", $"[{source}] ")

        Dim line As String = $"[{DateTime.Now:HH:mm:ss.fff}] [{elapsed.TotalSeconds:F3}s] [{levelStr}] {srcStr}{message}"

        SyncLock _lock
            Try
                File.AppendAllText(_logPath, line & Environment.NewLine)
            Catch
                ' Eșec silențios la scriere - nu vrem să blocăm aplicația
            End Try
        End SyncLock
    End Sub

    Private Shared Function SanitizeFileName(name As String) As String
        Dim invalid As Char() = Path.GetInvalidFileNameChars()
        Dim result As String = name
        For Each c In invalid
            result = result.Replace(c, "_"c)
        Next
        Return result
    End Function

    ''' <summary>
    ''' Returnează calea completă a fișierului de log (pentru debug).
    ''' </summary>
    Public Shared ReadOnly Property LogFilePath As String
        Get
            Return If(_logPath, "(neinițializat)")
        End Get
    End Property
End Class