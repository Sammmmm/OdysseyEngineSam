Attribute VB_Name = "modMap"
Option Explicit

Sub OpenMap(File As String)
    Dim TheMapData As String * 2677

    Open File For Binary As #15
    Get #15, 1, TheMapData
    Close #15

    ChDir App.Path
    CurDir App.Path

    LoadEditMap TheMapData
    RedrawMap = True
End Sub

Sub LoadEditMap(TheMapData As String)
    Dim A As Long, X As Long, Y As Long
    If Len(TheMapData) = 2677 Then
        With EditMap
            MapData = TheMapData
            MapDataLoadingArray() = StrConv(MapData, vbFromUnicode)
            .name = ClipString$(Mid$(MapData, 1, 30))
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1)) * 256 + Asc(Mid$(MapData, 36, 1))
            .MIDI = Asc(Mid$(MapData, 37, 1))
            .ExitUp = Asc(Mid$(MapData, 38, 1)) * 256 + Asc(Mid$(MapData, 39, 1))
            .ExitDown = Asc(Mid$(MapData, 40, 1)) * 256 + Asc(Mid$(MapData, 41, 1))
            .ExitLeft = Asc(Mid$(MapData, 42, 1)) * 256 + Asc(Mid$(MapData, 43, 1))
            .ExitRight = Asc(Mid$(MapData, 44, 1)) * 256 + Asc(Mid$(MapData, 45, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 46, 1)) * 256 + Asc(Mid$(MapData, 47, 1))
            .BootLocation.X = Asc(Mid$(MapData, 48, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 49, 1))
            .DeathLocation.Map = Asc(Mid$(MapData, 50, 1)) * 256 + Asc(Mid$(MapData, 51, 1))
            .DeathLocation.X = Asc(Mid$(MapData, 52, 1))
            .DeathLocation.Y = Asc(Mid$(MapData, 53, 1))
            .flags = Asc(Mid$(MapData, 54, 1))
            .Flags2 = Asc(Mid$(MapData, 55, 1))
            For A = 0 To 9    '56 - 86
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 56 + A * 3)) * 256 + Asc(Mid$(MapData, 57 + A * 3))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 58 + A * 3))
            Next A
            '86
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 86 + Y * 216 + X * 18
                        '1-10 = Tiles
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 9, 1))
                        .FGTile2 = Asc(Mid$(MapData, A + 10, 1)) * 256 + Asc(Mid$(MapData, A + 11, 1))
                        .Att = Asc(Mid$(MapData, A + 12, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 13, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 14, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 15, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 16, 1))
                        .Att2 = Asc(Mid$(MapData, A + 17, 1))
                    End With
                Next X
            Next Y

            If ExamineBit(.flags, 1) = True Then
                Indoors = True
            Else
                Indoors = False
            End If

            If ExamineBit(.flags, 2) = True Then
                AlwaysDark = True
            Else
                AlwaysDark = False
            End If

        End With
    End If
End Sub

Sub UploadMap()
    Dim TheMapData As String, St1 As String * 30
    Dim X As Long, Y As Long
    With EditMap
        If .Version < 2147483647 Then
            .Version = .Version + 1
        Else
            .Version = 1
        End If
        St1 = .name
        TheMapData = St1 + QuadChar(.Version) + DoubleChar$(CLng(.NPC)) + Chr$(.MIDI) + DoubleChar$(CLng(.ExitUp)) + DoubleChar$(CLng(.ExitDown)) + DoubleChar$(CLng(.ExitLeft)) + DoubleChar$(CLng(.ExitRight)) + DoubleChar(CLng(.BootLocation.Map)) + Chr$(.BootLocation.X) + Chr$(.BootLocation.Y) + DoubleChar$(CLng(.DeathLocation.Map)) + Chr$(.DeathLocation.X) + Chr$(.DeathLocation.Y)
        TheMapData = TheMapData + Chr$(.flags) + Chr$(.Flags2) + DoubleChar$(CLng(.MonsterSpawn(0).Monster)) + Chr$(.MonsterSpawn(0).Rate) + DoubleChar$(CLng(.MonsterSpawn(1).Monster)) + Chr$(.MonsterSpawn(1).Rate) + DoubleChar$(CLng(.MonsterSpawn(2).Monster)) + Chr$(.MonsterSpawn(2).Rate) + DoubleChar$(CLng(.MonsterSpawn(3).Monster)) + Chr$(.MonsterSpawn(3).Rate) + DoubleChar$(CLng(.MonsterSpawn(4).Monster)) + Chr$(.MonsterSpawn(4).Rate) + DoubleChar$(CLng(.MonsterSpawn(5).Monster)) + Chr$(.MonsterSpawn(5).Rate) + DoubleChar$(CLng(.MonsterSpawn(6).Monster)) + Chr$(.MonsterSpawn(6).Rate) + DoubleChar$(CLng(.MonsterSpawn(7).Monster)) + Chr$(.MonsterSpawn(7).Rate) + DoubleChar$(CLng(.MonsterSpawn(8).Monster)) + Chr$(.MonsterSpawn(8).Rate) + DoubleChar$(CLng(.MonsterSpawn(9).Monster)) + Chr$(.MonsterSpawn(9).Rate)
        For Y = 0 To 11
            For X = 0 To 11
                With .Tile(X, Y)
                    TheMapData = TheMapData + DoubleChar$(CLng(.Ground)) + DoubleChar$(CLng(.Ground2)) + DoubleChar$(CLng(.BGTile1)) + DoubleChar$(CLng(.BGTile2)) + DoubleChar$(CLng(.FGTile)) + DoubleChar$(CLng(.FGTile2)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(.Att2)
                End With
            Next X
        Next Y
    End With
    SendSocket Chr$(12) + TheMapData
End Sub

Sub LoadMapData(LoadMapData As String)
    On Error GoTo LoadError

    Dim A As Long, X As Long, Y As Long
    If Len(LoadMapData) = 2677 Then
        MapData = LoadMapData
        MapDataLoadingArray() = StrConv(MapData, vbFromUnicode)
        With Map
            .name = ClipString$(Mid$(MapData, 1, 30))
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1)) * 256 + Asc(Mid$(MapData, 36, 1))
            .MIDI = Asc(Mid$(MapData, 37, 1))
            .ExitUp = Asc(Mid$(MapData, 38, 1)) * 256 + Asc(Mid$(MapData, 39, 1))
            .ExitDown = Asc(Mid$(MapData, 40, 1)) * 256 + Asc(Mid$(MapData, 41, 1))
            .ExitLeft = Asc(Mid$(MapData, 42, 1)) * 256 + Asc(Mid$(MapData, 43, 1))
            .ExitRight = Asc(Mid$(MapData, 44, 1)) * 256 + Asc(Mid$(MapData, 45, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 46, 1)) * 256 + Asc(Mid$(MapData, 47, 1))
            .BootLocation.X = Asc(Mid$(MapData, 48, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 49, 1))
            .DeathLocation.Map = Asc(Mid$(MapData, 50, 1)) * 256 + Asc(Mid$(MapData, 51, 1))
            .DeathLocation.X = Asc(Mid$(MapData, 52, 1))
            .DeathLocation.Y = Asc(Mid$(MapData, 53, 1))
            .flags = Asc(Mid$(MapData, 54, 1))
            .Flags2 = Asc(Mid$(MapData, 55, 1))
            For A = 0 To 9    '56 - 86
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 56 + A * 3)) * 256 + Asc(Mid$(MapData, 57 + A * 3))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 58 + A * 3))
            Next A
            '86
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 86 + Y * 216 + X * 18
                        '1-10 = Tiles
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 9, 1))
                        .FGTile2 = Asc(Mid$(MapData, A + 10, 1)) * 256 + Asc(Mid$(MapData, A + 11, 1))
                        .Att = Asc(Mid$(MapData, A + 12, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 13, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 14, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 15, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 16, 1))
                        .Att2 = Asc(Mid$(MapData, A + 17, 1))
                    End With
                Next X
            Next Y
            
            If ExamineBit(.flags, 1) = True Then
                Indoors = True
            Else
                Indoors = False
            End If

            If ExamineBit(.flags, 2) = True Then
                AlwaysDark = True
            Else
                AlwaysDark = False
            End If

        End With
    End If

    Exit Sub

LoadError:
    PrintChat "Map failed to load - requesting again4", YELLOW
    RequestedMap = True
    SendSocket Chr$(45)
End Sub

Sub LoadMapFromCache(LoadMap As Long)
    ChDir App.Path
    CurDir App.Path

    On Error Resume Next
    Close #1
    On Error GoTo LoadError

    Open CacheDirectory + "/cache1.dat" For Random As #1 Len = 2677
    Get #1, LoadMap, MapData
    Close #1

    If Asc(Mid$(MapData, 1, 1)) > 0 Then
        Dim MapDataWorkingArray() As Byte
        MapDataWorkingArray() = StrConv(MapData, vbFromUnicode)
        EncryptDataString MapDataWorkingArray(0), CMap * CACHE_KEY Mod CACHE_KEY + 5
        MapData = StrConv(MapDataWorkingArray, vbUnicode)
    End If

    LoadMapData MapData

    Exit Sub

LoadError:
    PrintChat "Map failed to load - requesting again5", YELLOW
    RequestedMap = True
    SendSocket Chr$(45)
End Sub

Sub SaveMap(File As String)
    If Exists(File) Then
        If MsgBox("A map with that name already exists.  Saving will replace it.  Do you wish to continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
            On Error Resume Next
                Kill File
            On Error GoTo 0
        Else
            Exit Sub
        End If
    End If
    
    Dim TheMapData As String, St1 As String * 30
    Dim X As Long, Y As Long
    With Map
        St1 = .name
        TheMapData = St1 + QuadChar(.Version) + DoubleChar$(CLng(.NPC)) + Chr$(.MIDI) + DoubleChar$(CLng(.ExitUp)) + DoubleChar$(CLng(.ExitDown)) + DoubleChar$(CLng(.ExitLeft)) + DoubleChar$(CLng(.ExitRight)) + DoubleChar(CLng(.BootLocation.Map)) + Chr$(.BootLocation.X) + Chr$(.BootLocation.Y) + DoubleChar(CLng(.DeathLocation.Map)) + Chr$(.DeathLocation.X) + Chr$(.DeathLocation.Y)
        TheMapData = TheMapData + Chr$(.flags) + Chr$(.Flags2) + DoubleChar$(CLng(.MonsterSpawn(0).Monster)) + Chr$(.MonsterSpawn(0).Rate) + DoubleChar$(CLng(.MonsterSpawn(1).Monster)) + Chr$(.MonsterSpawn(1).Rate) + DoubleChar$(CLng(.MonsterSpawn(2).Monster)) + Chr$(.MonsterSpawn(2).Rate) + DoubleChar$(CLng(.MonsterSpawn(3).Monster)) + Chr$(.MonsterSpawn(3).Rate) + DoubleChar$(CLng(.MonsterSpawn(4).Monster)) + Chr$(.MonsterSpawn(4).Rate) + DoubleChar$(CLng(.MonsterSpawn(5).Monster)) + Chr$(.MonsterSpawn(5).Rate) + DoubleChar$(CLng(.MonsterSpawn(6).Monster)) + Chr$(.MonsterSpawn(6).Rate) + DoubleChar$(CLng(.MonsterSpawn(7).Monster)) + Chr$(.MonsterSpawn(7).Rate) + DoubleChar$(CLng(.MonsterSpawn(8).Monster)) + Chr$(.MonsterSpawn(8).Rate) + DoubleChar$(CLng(.MonsterSpawn(9).Monster)) + Chr$(.MonsterSpawn(9).Rate)
        For Y = 0 To 11
            For X = 0 To 11
                With .Tile(X, Y)
                    TheMapData = TheMapData + DoubleChar(CLng(.Ground)) + DoubleChar$(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + DoubleChar(CLng(.BGTile2)) + DoubleChar(CLng(.FGTile)) + DoubleChar(CLng(.FGTile2)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(.Att2)
                End With
            Next X
        Next Y
    End With

    Open File For Output As #1 Len = 2677
    Print #1, TheMapData
    Close #1
End Sub
