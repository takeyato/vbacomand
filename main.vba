Sub RunLongCommand()
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim longCommand As String
    longCommand = "ここに長いコマンドを入力してください" ' 例: "echo Part1 && echo Part2 && echo Part3 ..."

    Dim commands() As String
    commands = SplitCommand(longCommand, 256)
    
    Dim i As Integer
    For i = LBound(commands) To UBound(commands)
        wsh.Run "cmd /c " & commands(i), 0, True
    Next i
End Sub

Function SplitCommand(command As String, maxLength As Integer) As String()
    Dim parts() As String
    Dim part As String
    Dim i As Integer
    Dim startPos As Integer
    Dim endPos As Integer
    
    startPos = 1
    Do While startPos <= Len(command)
        endPos = startPos + maxLength - 1
        If endPos > Len(command) Then endPos = Len(command)
        
        part = Mid(command, startPos, endPos - startPos + 1)
        ReDim Preserve parts(i)
        parts(i) = part
        i = i + 1
        
        startPos = endPos + 1
    Loop
    
    SplitCommand = parts
End Function
