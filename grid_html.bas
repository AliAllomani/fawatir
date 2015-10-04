Attribute VB_Name = "grid_html"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (Destination As Any, Source As Any, ByVal Length As Long)



 Function ColorToRGB(ByVal lColor As Long) As String
   'Converts long value to HTML color string
   Dim lTemp As Long
   lTemp = lColor
   'VB colors are stored as BBGGRR
   'HTML colors are represented as RRGGBB
   'Flip the positions of the first and third bytes
   Call CopyMemory(ByVal VarPtr(lTemp) + 2, ByVal VarPtr(lColor), 1)
   Call CopyMemory(ByVal VarPtr(lTemp), ByVal VarPtr(lColor) + 2, 1)
   ColorToRGB = Right$("000000" & Hex(lTemp), 6) 'Pad with 0's
   ColorToRGB = "#" & ColorToRGB
End Function


Function Grid2HTML(grid As MSFlexGrid, OutputMethod As String, InHdr As String, Optional x_total As String)

Dim Meta$(10000)   '~10K Meta Data - reduce if you need the memory space.
Dim BigMeta$(10000) '~10K Building Meta Data
Dim YBuffer
Dim XBuffer

YBuffer = 0
XBuffer = 0

If UCase(Left(OutputMethod, 5)) = "FILE=" Then
    f = FreeFile
    Open Mid(OutputMethod, 6) For Output As f
    
    'Write your html header data here..
    If InHdr$ <> "%TAG%" And InHdr$ <> "" Then Print #f, InHdr$
    If InHdr$ = "%TAG%" Then Print #f, grid.Tag
    
    'Write table detail here.
    'Meta$(0) = Border Type
    'Meta$(1) = Font Style
    If grid.Appearance = flex3D Then Meta$(0) = "Border=" & Chr(34) & "2" & Chr(34)
    If grid.Appearance = flexFlat Then Meta$(0) = "Cellpadding=" & Chr(34) & "0" & Chr(34) & " Cellspacing=" & Chr(34) & "0" & Chr(34)
    
    Meta$(1) = "style=" & Chr(34) & "font-family: " & grid.Font.Name & "; font-size: " & grid.Font.Size + 4 & Chr(34)
    
    Meta$(2) = "bordercolor=" & Chr(34) & ColorToRGB(grid.GridColor) & Chr(34)
    
    meta_width = 0
    meta_height = 0
    
    '' edited here
    For xx = 1 To grid.Cols - 1
        meta_width = meta_width + grid.ColWidth(xx)
        For yy = 0 To grid.Rows - 1
            meta_height = meta_height + grid.RowHeight(yy)
        Next yy
    Next xx
    
    meta_height = Int(meta_height / Screen.TwipsPerPixelY) + YBuffer
    meta_width = Int(meta_width / Screen.TwipsPerPixelX) + XBuffer
    
    Meta$(3) = "width=" & Chr(34) & meta_width & Chr(34) '& " height=" & Chr(34) & meta_height & Chr(34)
    
    X = 0
    Do Until Meta$(X) = ""
    BigMeta$(1) = BigMeta$(1) & Meta$(X) & " "
    X = X + 1
    Loop
    
    Text$ = "<center><TABLE border=1" & Trim(BigMeta$(1)) & ">"
    Print #f, Text$ & vbCrLf
    For r = 0 To grid.Rows - 1
        grid.Row = r
        Print #f, "<tr >" & vbCrLf
        For c = 1 To grid.Cols - 1
        
        If grid.ColWidth(c) <> 0 Then
        'Create actual CELL
        grid.Col = c
        Text$ = grid.Text 'Get plain text
        If Text$ = "" Then Text$ = "&nbsp;"
        
        If grid.CellFontBold = True Or r < grid.FixedRows Then Text$ = "<B>" & Text$ & "</B>"
        If grid.CellFontItalic = True Then Text$ = "<I>" & Text$ & "</I>"
        If grid.CellFontUnderline = True Then Text$ = "<U>" & Text$ & "</U>"
        
        bg$ = ""
        If grid.CellBackColor > 0 Then bg$ = "Bgcolor=" & Chr(34) & ColorToRGB(Val(grid.CellBackColor)) & Chr(34)
        If c < grid.FixedCols Or r < grid.FixedRows And grid.BackColorFixed > 0 Then bg$ = "Bgcolor=" & Chr(34) & ColorToRGB(Val(grid.BackColorFixed)) & Chr(34)
            Print #f, "<td " & bg$ & " >" & Text$ & "</td>" & vbCrLf
        
        End If
        Next c
        
        Print #f, "</TR>" & vbCrLf
    Next r
    
    Print #f, "</table>" & vbCrLf
    
    If x_total <> "" Then
    Print #f, "<br><br><b>  «·„Ã„Ê⁄ «·ﬂ·Ì : </b>" & x_total & vbCrLf
    End If
    
    
    Close #f
End If
End Function

