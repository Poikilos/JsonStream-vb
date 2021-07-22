'
'MIT License 
'Copyright(c) 2021 Jake "Poikilos" Gustafson
'
'Permission Is hereby granted, free Of charge, to any person obtaining a copy
'of this software And associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, And/Or sell
'copies of the Software, And to permit persons to whom the Software Is
'furnished to do so, subject to the following conditions
'
'The above copyright notice And this permission notice shall be included In all
'copies Or substantial portions of the Software.
'
'THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or
'IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS Or COPYRIGHT HOLDERS BE LIABLE For ANY CLAIM, DAMAGES Or OTHER
'LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or OTHERWISE, ARISING FROM,
'OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS IN THE
'SOFTWARE.
'
Public Class JsonStream
    Private Shared Initialized As Boolean = False
    Private Shared Enclosures As Dictionary(Of Char, Char) = New Dictionary(Of Char, Char) '''< Store cononical JSON-defined end for each start of an enclosure.
    Private Const Openers As String = """{[" '''< Store cononical JSON-defined openers (keys of Enclosures as a string for faster access)
    Private Shared Escapes As Dictionary(Of String, String) = New Dictionary(Of String, String) '''< Store cononical JSON-defined escape sequences as keys where the value is the non-escaped equivalent.
    Private Const Escapees As String = "bfnrt""\" '''< Store cononical JSON-defined escape sequences as characters that must follow a backslash.
    ''' <summary>
    ''' This method initializes the Escapes Dictionary.
    ''' Call this at the start of any methods, even Shared ones
    ''' so that Escapes and Enclosures are populated first
    ''' (Since vb.net not allow Shared Sub New:
    ''' https://docs.microsoft.com/en-us/dotnet/visual-basic/misc/bc30480).
    ''' </summary>
    Private Shared Sub Init()
        If Initialized Then
            Exit Sub
        End If
        Initialized = True
        ' Set escape sequences used in the JSON standard (See
        ' <https://www.freeformatter.com/json-escape.html#:~:text=The%20following%20characters%20are%20reserved%20in%20JSON%20and,with%20%5C%22%207%20Backslash%20is%20replaced%20with%20%5C%5C>).
        Escapes("\b") = vbBack
        Escapes("\f") = vbFormFeed
        Escapes("\n") = vbLf
        Escapes("\r") = vbCr
        Escapes("\t") = vbTab
        Escapes("\""") = """"
        '^ Why no "\\"? Answer:
        '  The backslash is skipped since it must be added first to avoid doubling and removed
        '  last to avoid un-escaping And leaving letters behind.
        '^ Why no "\u"? Answer:
        '  Unicode sequences are not yet implemented.

        Enclosures(""""c) = """"c
        Enclosures("["c) = "]"c
        Enclosures("{"c) = "}"c

        'Ensure the integrity of the Openers cache:
        For i = 0 To (Openers.Length - 1)
            If Not Enclosures.ContainsKey(Openers(i)) Then
                Throw New MissingFieldException("Each character in Openers must exist as a key in Enclosures.")
            End If
        Next
        If Not Enclosures.Count > Openers.Length Then
            Throw New MissingFieldException("Each key in Enclosures must also be cached as a character in Openers.")
        End If
    End Sub
    ''' <summary>
    ''' Convert a vb string to JSON format (This does not add quotes).
    ''' </summary>
    ''' <param name="value">This must be a valid non-quoted string that is not Nothing.</param>
    ''' <returns>Get the string formatted and ready to be placed between quotes in a JSON file.</returns>
    Public Shared Function Escape(value As String) As String
        Init()
        If value Is Nothing Then
            Throw New ArgumentException("The argument should be a string.")
        End If
        ' See <https://stackoverflow.com/questions/18628917/how-can-iterate-in-dictionary-in-vb-net>:
        value = value.Replace("\", "\\")
        '^ The backslash must be added first so the backslashes in the escape sequences don't get doubled.
        For Each kvp As KeyValuePair(Of String, String) In Escapes
            value = value.Replace(kvp.Value, kvp.Key)
        Next kvp
        Return value
    End Function
    ''' <summary>
    ''' Remove the escape sequences from a JSON file.
    ''' </summary>
    ''' <param name="value">This must be a valid JSON-encoded string without enclosing quotes.</param>
    ''' <returns>Get a literal string ready to use in your VB code.</returns>
    Public Shared Function Unescape(value As String) As String
        Init()
        For Each kvp As KeyValuePair(Of String, String) In Escapes
            value = value.Replace(kvp.Key, kvp.Value)
        Next kvp
        value = value.Replace("\", "\\")
        '^ The backslash must be removed last so the backslashes in the escape sequences
        '  don 't get removed which would leave the escape code characters behind as literal characters.
        Return value
    End Function
    ''' <summary>
    ''' Split JSON by commas and colons that aren't enclosed by anything (by quotes nor brackets nor curly braces).
    ''' In other words, only split at the top level of the JSON content and leave subtrees intact.
    ''' </summary>
    ''' <param name="s">A string with one or more elements, not including the enclosures ("[...]" or "{...}")!</param>
    ''' <param name="IsDict">Allow splitting at the top level by ":" and "," rather than only ",".</param>
    ''' <param name="offset">Track the location in the file for better JSON syntax debugging output.</param>
    ''' <returns>Get a list of JSON elements that may contain sub-elements</returns>
    Public Shared Function SplitJson(s As String, IsDict As Boolean, offset As Integer) As List(Of String)
        Init()
        Dim chunks As List(Of String) = New List(Of String)
        Dim inEnclosures As String = "" '''< The current enclosures.
        Dim closing As Char = Nothing '''< This is the closing character for the last character in enclosures, as a cache to reduce Dictionary access.
        Dim escaper As Char = Nothing '''< This is the character that can prevent the current closing.
        Dim prevChar As Char = Nothing '''< The previous character.
        Dim chunkStart As Integer = 0
        Dim i As Integer = 0
        s = s.Trim()
        Dim splitters As String = ","
        If IsDict Then
            splitters = ":,"
        End If
        Dim prevControlChar = Nothing
        While i < s.Length
            Dim c As Char = s(i)
            If (c = closing) AndAlso (prevChar <> escaper) Then
                'End the quote or other enclosure if it is not escaped such as by "\"
                '(escaper is a separate variable to simplify the logic.
                'escaper Is only "\" when the closing is a quote).
                prevControlChar = c
                'If c = closing Then
                'If Not ((closing = """"c) AndAlso (prevChar = "\")) Then
                '^ If the end quote isn't escaped, then it closes the quote.
                'Since the subtree is closed, stop tracking the subtree by safely removing the enclosure:
                If inEnclosures.Length = 1 Then
                    inEnclosures = ""
                Else
                    inEnclosures = inEnclosures.Substring(0, inEnclosures.Length - 1)
                End If
                'Safely get the new closing (If there are none left, set closing to Nothing):
                If inEnclosures.Length > 0 Then
                    closing = Enclosures(inEnclosures(inEnclosures.Length - 1))
                Else
                    closing = Nothing
                End If
                If closing = """"c Then
                    'In JSON, If the enclosure is a quote then a backslash makes a quote inside
                    'of the quotes literal rather than ending the quote.
                    escaper = "\"
                End If

                'End If
                'End If
            ElseIf (closing <> """"c) AndAlso splitters.Contains(c) Then 'If not in quotes and c is a splitter
                'Save a chunk and step out of the chunk.
                If IsDict Then
                    If prevControlChar = c Then
                        Throw New FormatException("The JSON object at " & offset & " has a '" & c & "' at " & (offset + chunkStart - 1) & " but '" & splitters.Replace("" & c, "") & "' was expected.")
                    End If
                End If
                prevControlChar = c
                chunks.Add(s.Substring(chunkStart, i - chunkStart))
                chunkStart = i + 1  'Go past the comma since it isn't part of the next value.
                If chunkStart = s.Length Then ' Then
                    ' Since trim is done, the following condition isn't required: OrElse (s.Substring(chunkStart).Trim() = 0) Then
                    ' JSON objects must not end with a comma.
                    Throw New FormatException("The JSON object at " & offset & " ends at " & (offset + chunkStart - 1) & " with '" & c & "' but should end with an object.")
                End If
            ElseIf (closing <> """"c) AndAlso Openers.Contains(c) Then 'If not in quotes and c is an opener
                'Step into a new chunk.
                prevControlChar = c
                inEnclosures &= c
                closing = Enclosures(c)
                escaper = Nothing
                If closing = """"c Then
                    'In JSON, If the enclosure is a quote then a backslash makes a quote inside
                    'of the quotes literal rather than ending the quote.
                    escaper = "\"
                End If
            End If
            prevChar = c
        End While
        If chunkStart <> s.Length Then ' Then
            chunks.Add(s.Substring(chunkStart))
        End If
        Return chunks
    End Function
    ''' <summary>
    ''' Convert JSON syntax to a VB list or dictionary.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="offset">Track the location in the file for better JSON syntax debugging output. Start at 0.</param>
    ''' <returns></returns>
    Public Shared Function Deserialize(s As String, offset As Integer) As Object
        Init()
        s = s.Trim()
        If s.StartsWith("[") Then
            If Not s.EndsWith("]") Then
                Throw New FormatException("The JSON branch started with a ")
            End If
            Dim o As List(Of Object) = Nothing
        ElseIf s.StartsWith("{") Then
            Dim o As Dictionary(Of String, Object) = New Dictionary(Of String, Object)
            Dim chunks As List(Of String) = JsonStream.SplitJson(s.Substring(1, s.Length - 2), True, offset + 1)
        Else
            If offset <> 0 Then
                Throw New FormatException("JSON must be enclosed like ""[...]"" or ""{...}""")
            End If
        End If
    End Function
    Public Shared Function Serialize(o As Object) As String
        Init()
        'Try
        Dim s As String = o.Trim()
        Return Escape(s)
        'Catch
        'It is not a string if it doesn't have Trim.
        'End Try
    End Function

End Class
