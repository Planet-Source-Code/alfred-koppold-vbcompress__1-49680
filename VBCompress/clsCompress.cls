VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const OF_READ  As Long = &H0
Private Const OF_CREATE  As Long = &H1000
Private Const OFS_MAXPATHNAME   As Integer = 128
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const INFINITE As Long = -1
Private hSource As Long
Private ret As Long
Private hDestination As Long
Private bytesOfDest As Long
Private b As Single
Private m_SaveFilename As String
Private m_LoadFilename As String
Private Const BUFFERSIZE As Integer = 4096
Private Const N As Integer = 4096
Private Const F As Integer = 18
Private Const NIL As Long = N
Private Const THRESHOLD As Integer = 2
Private Infinity As Long
Private InBuffer() As Byte
Private OutBuffer() As Byte
Private DataLen As Long
Private BufferPointer As Long
Private textsize As Long
Private codesize As Long
Private textbuf(N + F - 1) As Byte
Private matchposition As Integer
Private matchlength As Byte
Private lson(N + 1) As Integer
Private rson(N + 257) As Integer
Private dad(N + 1) As Integer
Private Infile As Integer
Private Outfile As Integer
Private InLen As Long
Private InPointer As Long
Private OutPointer As Long
Private InFilePointer As Long
Private OutFilePointer As Long
Private Const Blocklen As Integer = 1
Private Type Header
    Magic As Long
    Magic2 As Long
    Magic3 As Integer
    Filesize As Long
End Type
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Compressheader As Header
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)
Private Declare Function LZCopy Lib "lz32.dll" (ByVal hfSource As Long, _
                                                ByVal hfDest As Long) As Long
Private Declare Function LZOpenFile Lib "lz32.dll" Alias "LZOpenFileA" (ByVal lpszFile As String, _
                                                                        lpOf As OFSTRUCT, _
                                                                        ByVal style As Long) As Long
Private Declare Sub LZClose Lib "lz32.dll" (ByVal hfFile As Long)

Public Function Compress(Optional LoadName As String = vbNullString, _
                         Optional SaveName As String = vbNullString) As Long

  Dim FileNumber  As Integer
  Dim FileNumber2 As Integer

    If LenB(LoadName) = 0 Then
        LoadName = m_LoadFilename
    End If
    If LenB(SaveName) = 0 Then
        SaveName = m_SaveFilename
    End If
    FileNumber = FreeFile
    Open LoadName For Binary As FileNumber
    FileNumber2 = FreeFile
    Open SaveName For Binary As FileNumber2
    LZCompFile FileNumber, FileNumber2
    Close FileNumber
    Close FileNumber2

End Function

Private Function DeGetByte() As Byte

    InFilePointer& = InFilePointer& + 1
    If InPointer& = BUFFERSIZE Then
        ReDim InBuffer(BUFFERSIZE - 1)
        Get Infile, , InBuffer
        BufferPointer& = 0
        DeGetByte = InBuffer(0)
        InPointer& = 1
     Else 'NOT INPOINTER&...
        DeGetByte = InBuffer(InPointer&)
        InPointer& = InPointer& + 1
    End If

End Function

Private Sub DePutByte(ByVal TheByte As Byte)

  Dim WriteBuffer() As Byte

    OutFilePointer& = OutFilePointer& + 1
    If OutPointer& = BUFFERSIZE - 1 Then
        OutBuffer(BUFFERSIZE - 1) = TheByte
        ReDim WriteBuffer(OutPointer)
        CopyMemory WriteBuffer(0), OutBuffer(0), OutPointer + 1
        Put Outfile, , WriteBuffer
        OutPointer& = 0
     Else 'NOT OUTPOINTER&...
        OutBuffer(OutPointer&) = TheByte
        OutPointer& = OutPointer& + 1
    End If

End Sub

Private Function EnGetByte() As Byte

    InFilePointer& = InFilePointer& + 1
    If InPointer& = BUFFERSIZE Then
        Get Infile, , InBuffer
        EnGetByte = InBuffer(0)
        InPointer& = 1
     Else 'NOT INPOINTER&...
        EnGetByte = InBuffer(InPointer&)
        InPointer& = InPointer& + 1
    End If

End Function

Private Sub EnPutByte(ByVal TheByte As Byte)

    OutFilePointer& = OutFilePointer& + 1
    If OutPointer& = BUFFERSIZE - 1 Then
        OutBuffer(BUFFERSIZE - 1) = TheByte
        BufferPointer& = 0
        Put Outfile, , OutBuffer
        OutPointer& = 0
     Else 'NOT OUTPOINTER&...
        OutBuffer(OutPointer&) = TheByte
        OutPointer& = OutPointer& + 1
    End If

End Sub

Public Function Expand(Optional LoadName As String = vbNullString, _
                       Optional SaveName As String = vbNullString) As Long

  Dim FileNumber  As Integer
  Dim FileNumber2 As Integer

    If LenB(LoadName) = 0 Then
        LoadName = m_LoadFilename
    End If
    If LenB(SaveName) = 0 Then
        SaveName = m_SaveFilename
    End If
    FileNumber = FreeFile
    Open LoadName For Binary As FileNumber
    FileNumber2 = FreeFile
    Open SaveName For Binary As FileNumber2
    LZDecompFile FileNumber, FileNumber2
    Close FileNumber
    Close FileNumber2

End Function

Public Function ExpandWithAPI(Optional LoadName As String = vbNullString, _
                              Optional SaveName As String = vbNullString) As Long

  Dim openStruct As OFSTRUCT

    If LenB(LoadName) = 0 Then
        LoadName = m_LoadFilename
    End If
    If LenB(SaveName) = 0 Then
        SaveName = m_SaveFilename
    End If
    hSource& = LZOpenFile(LoadName, openStruct, OF_READ)
    hDestination& = LZOpenFile(SaveName, openStruct, OF_CREATE)
    ret& = bytesOfDest = LZCopy(hSource&, hDestination&)
    LZClose hDestination&
    LZClose hSource&

End Function

Private Sub FDeleteNode(ByVal p As Integer)

  Dim q As Integer

    If dad(p) = NIL Then
        Exit Sub
    End If
    If rson(p) = NIL Then
        q = lson(p)
     ElseIf lson(p) = NIL Then 'NOT RSON(P)...
        q = rson(p)
     Else 'NOT LSON(P)...
        q = lson(p)
        If rson(q) <> NIL Then
            Do
                q = rson(q)
            Loop While rson(q) <> NIL
            rson(dad(q)) = lson(q)
            dad(lson(q)) = dad(q)
            lson(q) = lson(p)
            dad(lson(p)) = q
        End If
        rson(q) = rson(p)
        dad(rson(p)) = q
    End If
    dad(q) = dad(p)
    If rson(dad(p)) = p Then
        rson(dad(p)) = q
     Else 'NOT RSON(DAD(P))...
        lson(dad(p)) = q
    End If
    dad(p) = NIL

End Sub

Private Sub FInitTree()

  Dim i As Integer

    For i = N + 1 To (N + 256)
        rson(i) = NIL
    Next i
    For i = 0 To N - 1
        dad(i) = NIL
    Next i

End Sub

Private Sub FInsertNode(ByVal r As Integer)
  Dim i   As Integer
  Dim p   As Integer
  Dim cmp As Integer
  Dim key As Integer
  Dim x   As Long
  Dim x1  As Integer
  Dim x2  As Integer
    cmp = 1
    key = r
    p = N + 1 + textbuf(r)
    rson(r) = NIL
    lson(r) = NIL
    matchlength = 0
    Infinity& = 2147483647
    For x& = 0 To Infinity&
        If cmp >= 0 Then
            If rson(p) <> NIL Then
                p = rson(p)
             Else 'NOT RSON(P)...
                rson(p) = r
                dad(r) = p
                Exit Sub
            End If
         Else 'NOT CMP...
            If lson(p) <> NIL Then
                p = lson(p)
             Else 'NOT LSON(P)...
                lson(p) = r
                dad(r) = p
                Exit Sub
            End If
        End If
        For i = 1 To F - 1
            x1 = textbuf(r + i)
            x2 = textbuf(p + i)
            cmp = x1 - x2
            If cmp <> 0 Then
                Exit For
            End If
        Next i
        If i > matchlength Then
            matchposition = p
            matchlength = i
            If matchlength >= F Then
                Exit For
            End If
        End If
    Next x
    dad(r) = dad(p)
    lson(r) = lson(p)
    rson(r) = rson(p)
    dad(lson(p)) = r
    dad(rson(p)) = r
    If rson(dad(p)) = p Then
        rson(dad(p)) = r
     Else 'NOT RSON(DAD(P))...
        lson(dad(p)) = r
    End If
    dad(p) = NIL

End Sub

Private Sub FlushBuffer()

  Dim WriteBuffer() As Byte

    If OutPointer& > 0 Then
        ReDim WriteBuffer(OutPointer - 1)
        CopyMemory WriteBuffer(0), OutBuffer(0), OutPointer
        Put Outfile, , WriteBuffer
     Else 'NOT OUTPOINTER&...
    End If

End Sub

Public Property Get LoadFilename() As String

    LoadFilename = m_LoadFilename

End Property

Public Property Let LoadFilename(ByVal vNewValue As String)

    m_LoadFilename = vNewValue

End Property

Private Sub LZCompFile(ByVal file1 As Integer, _
                       ByVal file2 As Integer)

  Dim codebuf(17)     As Byte
  Dim codebufptr      As Integer
  Dim i               As Integer
  Dim cc              As Byte
  Dim r               As Integer
  Dim s               As Integer
  Dim lenn            As Integer
  Dim lastmatchlength As Integer
  Dim mask            As Byte

    Infile = file1
    Outfile = file2
    DataLen = LOF(Infile)
    With Compressheader
        .Filesize = DataLen
        .Magic = &H44445A53
        .Magic2 = &H3327F088
        .Magic3 = &H41
    End With 'Compressheader
    Put Outfile, , Compressheader
    FInitTree
    codebuf(0) = 0
    codebufptr = 1
    mask = 1
    s = 0
    r = N - F
    For i = s To r - 1
        textbuf(i) = 32
    Next i
    OpenBuffer
    For lenn = 0 To F - 1
        If InFilePointer& = InLen& Then
            Exit For
        End If
        cc = EnGetByte
        textbuf(r + lenn) = cc
    Next lenn
    textsize = lenn
    If textsize = 0 Then
        Exit Sub
    End If
    For i = 1 To F
        FInsertNode (r - i)
    Next i
    FInsertNode (r)
    Do
        If matchlength > lenn Then
            matchlength = lenn
        End If
        If matchlength <= THRESHOLD Then
            matchlength = 1
            codebuf(0) = (codebuf(0) Or mask)
            codebuf(codebufptr) = textbuf(r)
            codebufptr = codebufptr + 1
         Else 'NOT MATCHLENGTH...
            matchposition = matchposition + 2
            codebuf(codebufptr) = ((matchposition) And 255) 'Achtung normal ohne + 2
            codebufptr = codebufptr + 1
            codebuf(codebufptr) = (((matchposition \ 16) And 240) Or (matchlength - (THRESHOLD + 1))) And 255
            codebufptr = codebufptr + 1
        End If
        mask = (mask * 2) And 255
        If mask = 0 Then
            For i = 0 To codebufptr - 1
                EnPutByte codebuf(i)
            Next i
            codesize = codesize + codebufptr
            codebuf(0) = 0
            mask = 1
            codebufptr = mask
        End If
        lastmatchlength = matchlength
        For i = 0 To lastmatchlength - 1
            If InFilePointer& = InLen& Then
                Exit For
            End If
            cc = EnGetByte
            FDeleteNode (s)
            textbuf(s) = cc
            If s < (F - 1) Then
                textbuf(s + N) = cc
            End If
            s = (s + 1) And (N - 1)
            r = (r + 1) And (N - 1)
            FInsertNode (r)
        Next i
        textsize = textsize + i
        Do While i < lastmatchlength
            i = i + 1
            FDeleteNode (s)
            s = (s + 1) And (N - 1)
            r = (r + 1) And (N - 1)
            lenn = lenn - 1
            If lenn <> 0 Then
                FInsertNode (r)
            End If
        Loop
    Loop While lenn > 0
    If codebufptr > 1 Then
        For i = 0 To codebufptr - 1
            EnPutByte codebuf(i)
        Next i
        codesize = codesize + codebufptr
    End If
    FlushBuffer

End Sub

Private Sub LZDecompFile(ByVal file1 As Integer, _
                         ByVal file2 As Integer)

  Dim flags     As Byte
  Dim flagcount As Byte
  Dim x         As Long
  Dim bb        As Byte
  Dim b1        As Byte
  Dim b2        As Byte
  Dim i         As Integer
  Dim j         As Integer
  Dim k         As Integer
  Dim r         As Integer

    Infile = file1
    Outfile = file2
    DataLen = LOF(Infile) - 14
    Get Infile, , Compressheader
    For i = 0 To (N - F) - 1
        textbuf(i) = 32
    Next i
    r = N - F
    flags = 0
    flagcount = 8
    OpenBuffer
    Infinity& = 2147483647
    For x& = 0 To Infinity&
        If flagcount = 8 Then
            If InFilePointer& = InLen& Then
                Exit For
            End If
            bb = DeGetByte()
            flagcount = 0
            flags = bb
        End If
        If (flags And 1) = 1 Then
            If InFilePointer& = InLen& Then
                Exit For
            End If
            bb = DeGetByte()
            DePutByte bb
            textbuf(r) = bb
            r = r + 1
            r = (r And (N - 1))
         Else 'NOT (FLAGS...
            If InFilePointer& = InLen& Then
                Exit For
            End If
            b1 = DeGetByte()
            b1 = b1
            b! = b1
            If InFilePointer& = InLen& Then
                Exit For
            End If
            b2 = DeGetByte()
            i = b1
            j = b2
            i = i Or ((j And 240) * 16) 'Standort der Übereinstimmung
            i = i - 2 'Achtung nomal nicht - 2
            j = (j And 15) + THRESHOLD 'Länge der Übereinstimmung
            For k = 0 To j
                bb = textbuf((i + k) And (N - 1))
                DePutByte bb
                textbuf(r) = bb
                r = r + 1
                r = (r And (N - 1))
            Next k
        End If
        flags = (flags \ 2)
        flagcount = flagcount + 1
    Next x
    FlushBuffer

End Sub

Private Sub OpenBuffer()

    ReDim InBuffer(BUFFERSIZE - 1) As Byte
    ReDim OutBuffer(BUFFERSIZE - 1) As Byte
    InLen& = DataLen&
    Get Infile, , InBuffer
    InPointer& = 0
    InFilePointer& = 0
    OutPointer& = 0
    OutFilePointer& = 0

End Sub

Public Property Get SaveFilename() As String

    SaveFilename = m_SaveFilename

End Property

Public Property Let SaveFilename(ByVal vNewValue As String)

    m_SaveFilename = vNewValue

End Property
