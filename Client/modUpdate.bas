Attribute VB_Name = "modUpdate"
Option Explicit
Dim TileSource As RECT
    
Sub UpdateGame()
    Dim A As Long, B As Long, C As Long, D As Long
    Dim TempStr As String, TempVar As Byte
    
    For A = 1 To MaxUsers
        With Player(A)
            If .Map = CMap Then
                If Not .status = 25 Then
                    If .Sprite <= MaxSprite Then
                        'Move Player
                        If .XO < .X * 32 Then
                            D = .XO \ 16
                            .XO = (.X - 1) * 32 + ((timeGetTime - .WalkStart) / (3200 / .WalkStep)) * 32
                            If (.XO >= .X * 32) Then .XO = .X * 32
                            
                            If .XO \ 16 <> D Then .W = 1 - .W
                        ElseIf .XO > .X * 32 Then
                            D = .XO \ 16
                            .XO = (.X + 1) * 32 - ((timeGetTime - .WalkStart) / (3200 / .WalkStep)) * 32
                            If (.XO <= .X * 32) Then .XO = .X * 32
                            
                            If .XO \ 16 <> D Then .W = 1 - .W
                        End If
                        If .YO < .Y * 32 Then
                            D = .YO \ 16
                            .YO = (.Y - 1) * 32 + ((timeGetTime - .WalkStart) / (3200 / .WalkStep)) * 32
                            If (.YO >= .Y * 32) Then .YO = .Y * 32
                            
                            If .YO \ 16 <> D Then .W = 1 - .W
                        ElseIf .YO > .Y * 32 Then
                            D = .YO \ 16
                            .YO = (.Y + 1) * 32 - ((timeGetTime - .WalkStart) / (3200 / .WalkStep)) * 32
                            If (.YO <= .Y * 32) Then .YO = .Y * 32
                            
                            If .YO \ 16 <> D Then .W = 1 - .W
                        End If
                    End If
                End If
            End If
        End With
    Next A
        
    'Move you
    If CXO < CX * 32 Then
        'CXO = CXO + CWalkStep
        D = CXO \ 16
        CXO = (CX - 1) * 32 + ((timeGetTime - CWalkStart) / (TargetMoveTicks / CWalkStep)) * 32
        If CXO >= CX * 32 Then
            CXO = CX * 32
        End If
        
        If CXO \ 16 <> D Then
            CWalk = 1 - CWalk
            If CWalk = 0 Then PlayWav 4
        End If
    ElseIf CXO > CX * 32 Then
        'CXO = CXO - CWalkStep
        D = CXO \ 16
        CXO = (CX + 1) * 32 - ((timeGetTime - CWalkStart) / (TargetMoveTicks / CWalkStep)) * 32
        If CXO <= CX * 32 Then
            CXO = CX * 32
        End If
        
        If CXO \ 16 <> D Then
            CWalk = 1 - CWalk
            If CWalk = 0 Then PlayWav 4
        End If
    End If
    If CYO < CY * 32 Then
        'CYO = CYO + CWalkStep
        D = CYO \ 16
        CYO = (CY - 1) * 32 + ((timeGetTime - CWalkStart) / (TargetMoveTicks / CWalkStep)) * 32
        If CYO >= CY * 32 Then
            CYO = CY * 32
        End If
        
        If CYO \ 16 <> D Then
            CWalk = 1 - CWalk
            If CWalk = 0 Then PlayWav 4
        End If
    ElseIf CYO > CY * 32 Then
        'CYO = CYO - CWalkStep
        D = CYO \ 16
        CYO = (CY + 1) * 32 - ((timeGetTime - CWalkStart) / (TargetMoveTicks / CWalkStep)) * 32
        If CYO <= CY * 32 Then
            CYO = CY * 32
        End If
        
        If CYO \ 16 <> D Then
            CWalk = 1 - CWalk
            If CWalk = 0 Then PlayWav 4
        End If
    End If
    
    'Move monsters
    For A = 0 To MaxMonsters
        With Map.Monster(A)
            If .Monster > 0 Then
                C = Monster(.Monster).Sprite
                If C > 0 And C <= MaxSprite Then
                    If .XO < .X * 32 Then
                        D = .XO \ 16
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .XO = (.X - 1) * 32 + ((timeGetTime - .WalkStart) / (TargetMoveTicks / MonsterWalkSpeed)) * 32
                        Else
                            .XO = (.X - 1) * 32 + ((timeGetTime - .WalkStart) / (TargetMoveTicks / MonsterRunSpeed)) * 32
                        End If
                        If (.XO >= .X * 32) Then .XO = .X * 32
                        
                        If D <> .XO \ 16 Then .W = 1 - .W
                    ElseIf .XO > .X * 32 Then
                        D = .XO \ 16
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .XO = (.X + 1) * 32 - ((timeGetTime - .WalkStart) / (TargetMoveTicks / MonsterWalkSpeed)) * 32
                        Else
                            .XO = (.X + 1) * 32 - ((timeGetTime - .WalkStart) / (TargetMoveTicks / MonsterRunSpeed)) * 32
                        End If
                        If (.XO <= .X * 32) Then .XO = .X * 32
                        
                        If D <> .XO \ 16 Then .W = 1 - .W
                    End If
                    If .YO < .Y * 32 Then
                        D = .YO \ 16
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .YO = (.Y - 1) * 32 + ((timeGetTime - .WalkStart) / (TargetMoveTicks / MonsterWalkSpeed)) * 32
                        Else
                            .YO = (.Y - 1) * 32 + ((timeGetTime - .WalkStart) / (TargetMoveTicks / MonsterRunSpeed)) * 32
                        End If
                        If (.YO >= .Y * 32) Then .YO = .Y * 32
                        
                        If D <> .YO \ 16 Then .W = 1 - .W
                    ElseIf .YO > .Y * 32 Then
                        D = .YO \ 16
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .YO = (.Y + 1) * 32 - ((timeGetTime - .WalkStart) / (TargetMoveTicks / MonsterWalkSpeed)) * 32
                        Else
                            .YO = (.Y + 1) * 32 - ((timeGetTime - .WalkStart) / (TargetMoveTicks / MonsterRunSpeed)) * 32
                        End If
                        If (.YO <= .Y * 32) Then .YO = .Y * 32
                        
                        If D <> .YO \ 16 Then .W = 1 - .W
                    End If
                End If
            End If
        End With
    Next A
    
    
    'Update Projectiles
    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite > 0 Then
                Select Case .TargetType
                Case pttCharacter
                    If .TargetNum = Character.index Then
                        .X = CXO
                        .Y = CYO
                    Else
                        .X = Player(.TargetNum).XO
                        .Y = Player(.TargetNum).YO
                    End If

                    If Tick - .TimeStamp >= .speed Then
                        If .Frame < .TotalFrames Then
                            .Frame = .Frame + 1
                        Else
                            If .CurLoop = .LoopCount Then
                                If .EndSound > 0 Then
                                    PlayWav .EndSound
                                End If
                                DestroyEffect A
                            Else
                                .CurLoop = .CurLoop + 1
                                .Frame = 0
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttPlayer
                    If .TargetNum = Character.index Then
                        .TargetX = CXO
                        .TargetY = CYO
                    Else
                        .TargetX = Player(.TargetNum).XO
                        .TargetY = Player(.TargetNum).YO
                    End If
                    If .X < .TargetX Then .X = .X + 8
                    If .X > .TargetX Then .X = .X - 8
                    If .Y < .TargetY Then .Y = .Y + 8
                    If .Y > .TargetY Then .Y = .Y - 8

                    If Tick - .TimeStamp >= .speed Then
                        If .X = .TargetX Then
                            If .Y = .TargetY Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    DestroyEffect A
                                End If
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttMonster
                    .TargetX = Map.Monster(.TargetNum).XO
                    .TargetY = Map.Monster(.TargetNum).YO
                    If .X < .TargetX Then .X = .X + 8
                    If .X > .TargetX Then .X = .X - 8
                    If .Y < .TargetY Then .Y = .Y + 8
                    If .Y > .TargetY Then .Y = .Y - 8

                    If Tick - .TimeStamp >= .speed Then
                        If .X = .TargetX Then
                            If .Y = .TargetY Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    If .EndSound > 0 Then PlayWav .EndSound
                                    DestroyEffect A
                                End If
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttTile
                    If Tick - .TimeStamp >= .speed Then
                        If .Frame < .TotalFrames Then
                            .Frame = .Frame + 1
                        Else
                            If .CurLoop = .LoopCount Then
                                If .EndSound > 0 Then
                                    PlayWav .EndSound
                                End If
                                DestroyEffect (A)
                            Else
                                .CurLoop = .CurLoop + 1
                                .Frame = 0
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttProject
                    If (.X = .TargetX And .Y = .TargetY) Or (.X < -32 Or .Y < -16 Or .X > 12 * 32 + 16 Or .Y > 12 * 32 + 16) Then
                        If .TotalFrames > 0 Then
                            If .Frame < .TotalFrames Then
                                .Frame = .Frame + 1
                            Else
                                If .EndSound > 0 Then PlayWav .EndSound
                                DestroyEffect A
                            End If
                        Else
                            If .EndSound > 0 Then PlayWav .EndSound
                            DestroyEffect A
                        End If
                    Else
                        If .Direction = 3 Then .X = (.SourceX * 32) + ((timeGetTime - .StartTime) / (TargetMoveTicks / .speed)) * 32
                        If .Direction = 2 Then .X = (.SourceX * 32) - ((timeGetTime - .StartTime) / (TargetMoveTicks / .speed)) * 32
                        If .Direction = 1 Then .Y = (.SourceY * 32) + ((timeGetTime - .StartTime) / (TargetMoveTicks / .speed)) * 32
                        If .Direction = 0 Then .Y = (.SourceY * 32) - ((timeGetTime - .StartTime) / (TargetMoveTicks / .speed)) * 32
                        If .Alternate = True Then
                            Select Case .Type
                            Case 2
                                .offset = 1 - .offset
                                .Frame = .offset
                            Case 4
                                If .offset = 3 Then .offset = 0 Else .offset = .offset + 1
                                .Frame = .offset
                            End Select
                        End If
                        C = (.X / 32)
                        D = (.Y / 32)
                        
                        If (C >= 0 And C <= 11 And D >= 0 And D <= 11) Then
                            'Projectile Collision
                            Select Case Map.Tile(C, D).Att
                            Case 1, 2, 3, 14, 16
                                .TargetX = .X
                                .TargetY = .Y
                            Case 19    'Light
                                If ExamineBit(Map.Tile(C, D).AttData(2), 0) = 1 Then
                                    .TargetX = .X
                                    .TargetY = .Y
                                End If
                            Case 20    'Light Dampening
                                If ExamineBit(Map.Tile(C, D).AttData(3), 0) Then
                                    .TargetX = .X
                                    .TargetY = .Y
                                End If
                            End Select
                            Select Case Map.Tile(C, D).Att2
                            Case 1, 14, 16
                                .TargetX = .X
                                .TargetY = .Y
                            End Select
                            
                            If NoDirectionalWalls(CByte(.X / 32), CByte(.Y / 32), .Direction) = False Then
                                .TargetX = .X
                                .TargetY = .Y
                            End If
                        End If
                        For B = 0 To MaxMonsters
                            If Map.Monster(B).X = C Then
                                If Map.Monster(B).Y = D Then
                                    If Map.Monster(B).Monster > 0 Then
                                        If .Creator = Character.index Then
                                            If .Damage > 0 Then
                                                TempVar = (CMap + CX + CY) Mod 250
                                                If .Magic > 0 Then
                                                    'Magic Projectile
                                                    TempStr = Chr$(TempVar) + Chr$(1) + Chr$(B) + Chr$(.Damage)
                                                    SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                Else
                                                    'Normal Projectile
                                                    TempStr = Chr$(TempVar) + Chr$(2) + Chr$(B) + Chr$(.Damage)
                                                    SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                End If
                                            Else
                                                SendSocket Chr$(73) & Chr$(B)
                                            End If
                                        End If

                                        .TargetX = .X
                                        .TargetY = .Y
                                    End If
                                End If
                            End If
                        Next B
                        For B = 1 To MaxUsers
                            If Player(B).X = C Then
                                If Player(B).Y = D Then
                                    If Player(B).Map = CMap Then
                                        If Not B = .Creator Then
                                            If Player(B).IsDead = False Then
                                                Dim Collide As Boolean
                                                If Character.Guild > 0 Then
                                                    If Player(B).Guild = 0 Then
                                                        If ExamineBit(Map.flags, 0) = False And ExamineBit(Map.flags, 6) = False Then

                                                        Else
                                                            Collide = True
                                                        End If
                                                    Else
                                                        Collide = True
                                                    End If
                                                Else
                                                    Collide = True
                                                End If
                                                If Collide = True Then
                                                    .TargetX = .X
                                                    .TargetY = .Y
                                                    If .Creator = Character.index Then
                                                        If .Damage > 0 Then
                                                            TempVar = CMap Mod 250

                                                            If .Magic > 0 Then
                                                                'Magic Projectile
                                                                TempStr = Chr$(TempVar) + Chr$(3) + Chr$(B) + Chr$(.Damage)
                                                                SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                                Exit For
                                                            Else
                                                                'Normal Projectile
                                                                TempStr = Chr$(TempVar) + Chr$(4) + Chr$(B) + Chr$(.Damage)
                                                                SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                                Exit For
                                                            End If
                                                        Else
                                                            SendSocket Chr$(74) & Chr$(B)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next B
                        If CX = C Then
                            If CY = D Then
                                If Not .Creator = Character.index Then
                                    .TargetX = .X
                                    .TargetY = .Y
                                End If
                            End If
                        End If
                    End If
                    .TimeStamp = Tick
                End Select
            End If
        End With
    Next A
End Sub

Sub DrawNextFrame()
    Dim A As Long, B As Long, C As Long, E As Double, r As RECT
    Dim TempVar As Byte, TempStr As String

    If RestoreDirectDraw = True Then
        On Error Resume Next
        UnloadDirectDraw
        InitDirectDraw
        LoadSurfaces
        RedrawMap = True
        On Error GoTo 0
        RestoreDirectDraw = False
    End If

    If RedrawMap = True Then
        DrawMap
        RedrawMap = False
    End If

    'Copy Back Buffer to Viewport
    If CurFrame = 0 Then
        BackBufferSurf.BltFast 0, 0, BGTile1Buffer, FullMapRect, DDBLTFAST_WAIT
    Else
        BackBufferSurf.BltFast 0, 0, BGTile2Buffer, FullMapRect, DDBLTFAST_WAIT
    End If

    For A = 1 To MaxUsers
        With Player(A)
            If .Map = CMap Then
                If Not .status = 25 Then
                    If .Sprite <= MaxSprite Then
                        'Draw Player
                        If .A > 0 Then
                            B = .D * 3 + 2
                            .A = .A - 1
                        Else
                            B = .D * 3 + .W
                        End If

                        If Player(A).IsDead Then
                            Draw .XO, .YO, 32, 32, DDSTiles, (623 Mod 7) * 32, (623 / 7) * 32, True
                        Else
                            If (Map.Tile(.X, .Y).Att = 25 And ExamineBit(Map.Tile(.X, .Y).AttData(1), 0) = True) Then 'sprite half tile
                                Draw .XO, .YO - 16, 32, 32 - Map.Tile(.X, .Y).AttData(0), DDSSprites, B * 32, (.Sprite - 1) * 32, True
                            Else
                                Draw .XO, .YO - 16, 32, 32, DDSSprites, B * 32, (.Sprite - 1) * 32, True
                            End If
                            If Player(A).HP > 0 Then
                                If Not Player(A).HP = Player(A).MaxHP Then
                                    Draw .XO + 3, .YO - 16, 2, 26, DDSHPBar, 0, 4, False
                                    Draw .XO + 3, .YO - 16, 2, 26 - (Player(A).HP / Player(A).MaxHP) * 26, DDSHPBar, 2, 4, False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next A

    'Draw You
    If CAttack > 0 Then
        B = CDir * 3 + 2
        CAttack = CAttack - 1
    Else
        B = CDir * 3 + CWalk
    End If

    If Character.IsDead = True Then
        Draw CXO, CYO, 32, 32, DDSTiles, (623 Mod 7) * 32, (623 / 7) * 32, True
    Else
        If (Map.Tile(CX, CY).Att = 25) And ExamineBit(Map.Tile(CX, CY).AttData(1), 0) = True Then 'sprite half tile
            Draw CXO, CYO - 16, 32, 32 - Map.Tile(CX, CY).AttData(0), DDSSprites, B * 32, (Character.Sprite - 1) * 32, True
        Else
            Draw CXO, CYO - 16, 32, 32, DDSSprites, B * 32, (Character.Sprite - 1) * 32, True
        End If
    End If

    For A = 0 To MaxMonsters
        With Map.Monster(A)
            If .Monster > 0 Then
                C = Monster(.Monster).Sprite
                If C > 0 And C <= MaxSprite Then
                    'Draw Monster
                    If .A > 0 Then
                        B = .D * 3 + 2
                        .A = .A - 1
                    Else
                        B = .D * 3 + .W
                    End If
                    If (Map.Tile(.X, .Y).Att = 25 And ExamineBit(Map.Tile(.X, .Y).AttData(1), 1) = True) Then 'sprite half tile
                        Draw .XO, .YO - 16, 32, 32 - Map.Tile(.X, .Y).AttData(0), DDSSprites, B * 32, (C - 1) * 32, True
                    Else
                        Draw .XO, .YO - 16, 32, 32, DDSSprites, B * 32, (C - 1) * 32, True
                    End If
                    If .HPBar = True Or Character.Access > 0 Then
                        Draw .XO + 3, .YO - 20, 26, 2, DDSHPBar, 0, 0, False
                        E = (.Life / Monster(.Monster).MaxLife)
                        If E > 1 Then E = 1
                        Draw .XO + 3, .YO - 20, E * 26, 2, DDSHPBar, 0, 2, False
                    End If
                End If
            End If
        End With
    Next A

    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite > 0 Then
                Draw .X, .Y - 16, 32, 32, DDSEffects, .Frame * 32, (.Sprite - 1) * 32, True
            End If
        End With
    Next A

    If CurFrame = 0 Then
        Call BackBufferSurf.BltFast(0, 0, FGTileBuffer, FullMapRect, DDBLTFAST_SRCCOLORKEY)
    Else
        Call BackBufferSurf.BltFast(0, 0, FGTile2Buffer, FullMapRect, DDBLTFAST_SRCCOLORKEY)
    End If

    Dim hdcBuffer As Long
    hdcBuffer = BackBufferSurf.GetDC
    SetBkMode hdcBuffer, Transparent

    With r
        .Left = CXO - 32
        .Right = CXO + 64
        .Top = CYO - 32
        .Bottom = CYO - 16
    End With

    If Character.Guild > 0 Then
        If Character.status = 1 And CurFrame = 0 Then
            Draw3dText hdcBuffer, r, Character.name, QBColor(4), 2
        Else
            If Character.status > 1 Then
                Draw3dText hdcBuffer, r, Character.name, StatusColors(Character.status), 2
            Else
                Draw3dText hdcBuffer, r, Character.name, QBColor(11), 2
            End If
        End If
    Else
        If Character.status = 2 Then
            Draw3dText hdcBuffer, r, Character.name, QBColor(14), 2
        ElseIf Character.status = 3 Then
            Draw3dText hdcBuffer, r, Character.name, QBColor(9), 2
        ElseIf Character.status = 1 And CurFrame = 0 Then
            Draw3dText hdcBuffer, r, Character.name, QBColor(4), 2
        Else
            Draw3dText hdcBuffer, r, Character.name, StatusColors(Character.status), 2
        End If
    End If

    If Character.status = 24 Then    'Rainbow
        Draw3dText hdcBuffer, r, Character.name, StatusColors((Int(Rnd * 23))), 2
    End If

    For A = 1 To MaxUsers
        With Player(A)
            If .Map = CMap Then
                If .IsDead = False Then
                    If .status = 9 Or .status = 25 Then

                    Else
                        r.Left = .XO - 32
                        r.Right = .XO + 64
                        r.Top = .YO - 32
                        r.Bottom = .YO - 16

                        If .status = 1 And CurFrame = 0 Then
                            Draw3dText hdcBuffer, r, .name, QBColor(4), 2
                        ElseIf .status = 1 And CurFrame = 1 Then
                            Draw3dText hdcBuffer, r, .name, QBColor(.Color), 2
                        ElseIf .status = 0 Then
                            Draw3dText hdcBuffer, r, .name, QBColor(.Color), 2
                        Else
                            If .status = 24 Then  'Rainbow
                                Draw3dText hdcBuffer, r, .name, StatusColors((Int(Rnd * 23))), 2
                            Else
                                Draw3dText hdcBuffer, r, .name, StatusColors(.status), 2
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next A

    For A = 1 To MaxFloatText    'Floating Text
        With FloatText(A)
            If .InUse = True Then
                With r
                    .Left = FloatText(A).X * 32 - 32
                    .Right = FloatText(A).X * 32 + 64
                    .Top = FloatText(A).Y * 32 - 32 + FloatText(A).FloatY
                    .Bottom = FloatText(A).Y * 32 - 16
                End With
                Draw3dText hdcBuffer, r, .Text, QBColor(.Color), 2
                If .Static = False Then
                    .FloatY = .FloatY - 1
                    If .FloatY <= -38 Then ClearFloatText CByte(A)
                End If
            End If
        End With
    Next A

    BackBufferSurf.ReleaseDC hdcBuffer

    If options.DisableLighting = False Then
        UpdateLights

        If ShadeMapFrame = 0 Then
            UpdateLightMap Lighting(0)
            ShadeMapFrame = options.LightingQuality
        Else
            ShadeMapFrame = ShadeMapFrame - 1
        End If

        BackBufferSurf.Lock EmptyRect, DDSDBackBuffer, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0

        BackBufferSurf.GetLockedArray ddsBufferArray()

        If Indoors = False Then
            If ExamineBit(Map.Flags2, 0) = True Then
                If options.Bit32 = True Then
                    Rain32 ddsBufferArray(0, 0), Tick
                Else
                    Rain16 ddsBufferArray(0, 0), Tick
                End If
            End If
            If ExamineBit(Map.Flags2, 1) = True Then
                If options.Bit32 = True Then
                    Snow32 ddsBufferArray(0, 0), Tick
                Else
                    Snow16 ddsBufferArray(0, 0), Tick
                End If
            End If
        End If

        If options.Bit32 = True Then
            ShadeMap32 ddsBufferArray(0, 0)
        Else
            ShadeMap16 ddsBufferArray(0, 0)
        End If

        BackBufferSurf.Unlock EmptyRect
    End If
    
    hdcBuffer = BackBufferSurf.GetDC
    SetBkMode hdcBuffer, Transparent
    
    DrawChatString hdcBuffer
    DrawInfoText hdcBuffer

    BackBufferSurf.ReleaseDC hdcBuffer

    If frmMain_Showing = False Then
        frmMain.Show
        RefreshInventory
        frmMain_Showing = True
    End If

    Call DX7.GetWindowRect(frmMain.picViewport.hwnd, MapRect)

    LastDrawReturn = -1
    On Error Resume Next
    LastDrawReturn = PrimarySurf.Blt(MapRect, BackBufferSurf, FullMapRect, DDBLT_WAIT)
    On Error GoTo 0

    If (LastDrawReturn <> 0) Then

        On Error Resume Next
        RestoreSurfaces

        LastDrawReturn = -1
        LastDrawReturn = PrimarySurf.Blt(MapRect, BackBufferSurf, FullMapRect, DDBLT_WAIT)

        On Error GoTo 0

        If LastDrawReturn <> 0 Then
            RestoreDirectDraw = True
        End If
    End If

End Sub

Sub RedrawMapTile(X As Byte, Y As Byte)
    If frmMain.Visible = False Then Exit Sub

    Dim TileSource As RECT
    Dim A As Long, B As Long
    If X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        TileSource.Left = X * 32
        TileSource.Top = Y * 32
        TileSource.Right = TileSource.Left + 32
        TileSource.Bottom = TileSource.Top + 32
        Call BGTile1Buffer.BltColorFill(TileSource, RGB(0, 0, 0))
        Call BGTile2Buffer.BltColorFill(TileSource, RGB(0, 0, 0))
        Call FGTileBuffer.BltColorFill(TileSource, RGB(0, 0, 0))
        Call FGTile2Buffer.BltColorFill(TileSource, RGB(0, 0, 0))
        If MapEdit = False Then
           DrawMapTile Map, X, Y
        Else
           DrawMapTile EditMap, X, Y
        End If
        For A = 0 To MaxMapObjects
            With Map.Object(A)
                If .Object > 0 And .X = X And .Y = Y Then
                    B = Object(.Object).Picture
                    If B > 0 Then
                        DrawObject BGTile1Buffer, B, .X, .Y
                        DrawObject BGTile2Buffer, B, .X, .Y
                    End If
                End If
            End With
        Next A
    End If
End Sub

Sub RedrawTile()
    BitBlt frmMain.picTile.hDC, 0, 0, 32, 32, 0, 0, 0, BLACKNESS

    If EditMode < 6 Then
        If CurTile > 0 Then
            DrawToDC 0, 0, 32, 32, frmMain.picTile.hDC, DDSTiles, ((CurTile - 1) Mod 7) * 32, Int((CurTile - 1) / 7) * 32
        End If
    Else
        If CurAtt > 0 Then
            DrawToDC 0, 0, 32, 32, frmMain.picTile.hDC, DDSAtts, ((CurAtt - 1) Mod 7) * 32, Int((CurAtt - 1) / 7) * 32
        End If
    End If
    frmMain.picTile.Refresh
End Sub
Sub RedrawTiles()
    BitBlt frmMain.picTiles.hDC, 0, 0, 224, 192, 0, 0, 0, BLACKNESS

    If EditMode < 6 Then
        DrawToDC 0, 0, 224, 192, frmMain.picTiles.hDC, DDSTiles, 0, CInt(TopY)
    Else
        DrawToDC 0, 0, 224, 192, frmMain.picTiles.hDC, DDSAtts, 0, 0
    End If
    frmMain.picTiles.Refresh
End Sub

Sub Draw3dText(DC As Long, TargetRect As RECT, St As String, lngColor As Long, Height As Integer)
    Dim ShadowRect As RECT
    With ShadowRect
        .Top = TargetRect.Top + Height
        .Left = TargetRect.Left + Height
        .Bottom = TargetRect.Bottom + Height
        .Right = TargetRect.Right + Height
    End With
    SetTextColor DC, RGB(10, 10, 10)
    DrawText DC, St, Len(St), ShadowRect, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK
    SetTextColor DC, lngColor
    DrawText DC, St, Len(St), TargetRect, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK
End Sub

Sub DrawMap()
    Dim A As Long, B As Long, X As Byte, Y As Byte
    Call BGTile1Buffer.BltColorFill(FullMapRect, RGB(0, 0, 0))
    Call BGTile2Buffer.BltColorFill(FullMapRect, RGB(0, 0, 0))
    Call FGTileBuffer.BltColorFill(FullMapRect, RGB(0, 0, 0))
    Call FGTile2Buffer.BltColorFill(FullMapRect, RGB(0, 0, 0))
    
    If CMap = 0 Then Exit Sub

    If options.DisableLighting = False Then
        ClearMapLights

        If ExamineBit(Map.Flags2, 0) = True Then 'Raining
            InitRain Tick
        End If

        If ExamineBit(Map.Flags2, 1) = True Then 'Snowing
            InitSnow Tick
        End If
    End If

    If MapEdit = False Then
        For X = 0 To 11
            For Y = 0 To 11
                DrawMapTile Map, X, Y
            Next Y
        Next X
    Else
        For X = 0 To 11
            For Y = 0 To 11
                DrawMapTile EditMap, X, Y
            Next Y
        Next X
    End If

    For A = 0 To MaxMapObjects
        With Map.Object(A)
            If .Object > 0 Then
                B = Object(.Object).Picture
                If B > 0 Then
                    DrawObject BGTile1Buffer, B, .X, .Y
                    DrawObject BGTile2Buffer, B, .X, .Y
                End If
            End If
        End With
    Next A

    If ExamineBit(Map.flags, 1) = True Then
        Indoors = True
    Else
        Indoors = False
    End If

    If ExamineBit(Map.flags, 2) = True Then
        AlwaysDark = True
    Else
        AlwaysDark = False
    End If

    If options.DisableLighting = False Then
        UpdateLights

        CreateLightMap Lighting(0), Darkness, MapDataLoadingArray(0), OutdoorLight
    End If
End Sub

Sub DrawMapTile(Map As MapData, X As Byte, Y As Byte)
    With Map.Tile(X, Y)
        If .Ground > 0 Then
            DrawTile BGTile1Buffer, .Ground, X, Y
            DrawTile BGTile2Buffer, .Ground, X, Y
        End If
        If .Ground2 > 0 Then
            DrawTile BGTile1Buffer, .Ground2, X, Y
            DrawTile BGTile2Buffer, .Ground2, X, Y
        End If
        If .BGTile1 > 0 Then
            DrawTile BGTile1Buffer, .BGTile1, X, Y
        End If
        If .BGTile2 > 0 Then
            DrawTile BGTile2Buffer, .BGTile2, X, Y
        ElseIf .BGTile1 > 0 Then
            DrawTile BGTile2Buffer, .BGTile1, X, Y
        End If
        If .FGTile > 0 Then
            If .Att = 23 Then
                If (ExamineBit(.AttData(0), 4) = False) Then 'bg checked tiles are visible
                    If (ExamineBit(.AttData(0), 0) = True) Then DrawTile BGTile1Buffer, .FGTile, X, Y, 16, 0, 0, 16 'upper right
                    If (ExamineBit(.AttData(0), 1) = True) Then DrawTile BGTile1Buffer, .FGTile, X, Y, 0, 16, 0, 16 'upper left
                    If (ExamineBit(.AttData(0), 2) = True) Then DrawTile BGTile1Buffer, .FGTile, X, Y, 0, 16, 16, 0 'lower left
                    If (ExamineBit(.AttData(0), 3) = True) Then DrawTile BGTile1Buffer, .FGTile, X, Y, 16, 0, 16, 0 'lower right
                End If
                
                If (ExamineBit(.AttData(0), 5) = False) Then 'unchecked checked tiles are visible
                    If (ExamineBit(.AttData(0), 0) = False) Then DrawTile FGTileBuffer, .FGTile, X, Y, 16, 0, 0, 16 'upper right
                    If (ExamineBit(.AttData(0), 1) = False) Then DrawTile FGTileBuffer, .FGTile, X, Y, 0, 16, 0, 16 'upper left
                    If (ExamineBit(.AttData(0), 2) = False) Then DrawTile FGTileBuffer, .FGTile, X, Y, 0, 16, 16, 0 'lower left
                    If (ExamineBit(.AttData(0), 3) = False) Then DrawTile FGTileBuffer, .FGTile, X, Y, 16, 0, 16, 0 'lower right
                End If
            Else
                DrawTile FGTileBuffer, .FGTile, X, Y
            End If
        End If
        
        If .FGTile2 > 0 Then
            If .Att = 23 Then
                If (ExamineBit(.AttData(0), 4) = False) Then 'bg checked tiles are visible
                    If (ExamineBit(.AttData(0), 0) = True) Then DrawTile BGTile2Buffer, .FGTile2, X, Y, 16, 0, 0, 16 'upper right
                    If (ExamineBit(.AttData(0), 1) = True) Then DrawTile BGTile2Buffer, .FGTile2, X, Y, 0, 16, 0, 16 'upper left
                    If (ExamineBit(.AttData(0), 2) = True) Then DrawTile BGTile2Buffer, .FGTile2, X, Y, 0, 16, 16, 0 'lower left
                    If (ExamineBit(.AttData(0), 3) = True) Then DrawTile BGTile2Buffer, .FGTile2, X, Y, 16, 0, 16, 0 'lower right
                End If
                
                If (ExamineBit(.AttData(0), 5) = False) Then 'unchecked checked tiles are visible
                    If (ExamineBit(.AttData(0), 0) = False) Then DrawTile FGTile2Buffer, .FGTile2, X, Y, 16, 0, 0, 16 'upper right
                    If (ExamineBit(.AttData(0), 1) = False) Then DrawTile FGTile2Buffer, .FGTile2, X, Y, 0, 16, 0, 16 'upper left
                    If (ExamineBit(.AttData(0), 2) = False) Then DrawTile FGTile2Buffer, .FGTile2, X, Y, 0, 16, 16, 0 'lower left
                    If (ExamineBit(.AttData(0), 3) = False) Then DrawTile FGTile2Buffer, .FGTile2, X, Y, 16, 0, 16, 0 'lower right
                End If
            Else
                DrawTile FGTile2Buffer, .FGTile2, X, Y
            End If
        ElseIf .FGTile > 0 Then
            If .Att = 23 Then
                If (ExamineBit(.AttData(0), 4) = False) Then 'bg checked tiles are visible
                    If (ExamineBit(.AttData(0), 0) = True) Then DrawTile BGTile2Buffer, .FGTile, X, Y, 16, 0, 0, 16 'upper right
                    If (ExamineBit(.AttData(0), 1) = True) Then DrawTile BGTile2Buffer, .FGTile, X, Y, 0, 16, 0, 16 'upper left
                    If (ExamineBit(.AttData(0), 2) = True) Then DrawTile BGTile2Buffer, .FGTile, X, Y, 0, 16, 16, 0 'lower left
                    If (ExamineBit(.AttData(0), 3) = True) Then DrawTile BGTile2Buffer, .FGTile, X, Y, 16, 0, 16, 0 'lower right
                End If
                
                If (ExamineBit(.AttData(0), 5) = False) Then 'unchecked checked tiles are visible
                    If (ExamineBit(.AttData(0), 0) = False) Then DrawTile FGTile2Buffer, .FGTile, X, Y, 16, 0, 0, 16 'upper right
                    If (ExamineBit(.AttData(0), 1) = False) Then DrawTile FGTile2Buffer, .FGTile, X, Y, 0, 16, 0, 16 'upper left
                    If (ExamineBit(.AttData(0), 2) = False) Then DrawTile FGTile2Buffer, .FGTile, X, Y, 0, 16, 16, 0 'lower left
                    If (ExamineBit(.AttData(0), 3) = False) Then DrawTile FGTile2Buffer, .FGTile, X, Y, 16, 0, 16, 0 'lower right
                End If
            Else
                DrawTile FGTile2Buffer, .FGTile, X, Y
            End If
        End If
        
        If EditMode >= 6 And MapEdit Then
            If .Att2 > 0 Then
                DrawAtt FGTileBuffer, .Att2, X, Y, 12, 12
                DrawAtt FGTile2Buffer, .Att2, X, Y, 12, 12
            End If
            If .Att > 0 Then
                DrawAtt FGTileBuffer, .Att, X, Y, 4, 4
                DrawAtt FGTile2Buffer, .Att, X, Y, 4, 4
            End If
        End If
        If .Att = 19 Then
            If options.DisableLighting = False Then AddMapLight X * 32 + 16, Y * 32 + 16, .AttData(0), .AttData(1)
        End If
    End With
End Sub

Sub DrawTile(ByRef Surface As DirectDrawSurface4, tileIndex As Integer, X As Byte, Y As Byte, _
    Optional trimLeft As Long = 0, Optional trimRight As Long = 0, Optional trimTop As Long = 0, Optional trimBottom As Long = 0)
    
    TileSource.Left = (((tileIndex - 1) Mod 7) * 32) + trimLeft
    TileSource.Top = (Int((tileIndex - 1) / 7) * 32) + trimTop
    TileSource.Right = (((tileIndex - 1) Mod 7) * 32) + 32 - trimRight
    TileSource.Bottom = (Int((tileIndex - 1) / 7) * 32) + 32 - trimBottom
    Call Surface.BltFast(X * 32 + trimLeft, Y * 32 + trimTop, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
End Sub

Sub DrawObject(ByRef Surface As DirectDrawSurface4, objectIndex As Long, X As Byte, Y As Byte)
    TileSource.Left = 0
    TileSource.Top = (objectIndex - 1) * 32
    TileSource.Right = TileSource.Left + 32
    TileSource.Bottom = TileSource.Top + 32
    Call Surface.BltFast(X * 32, Y * 32, DDSObjects, TileSource, DDBLTFAST_SRCCOLORKEY)
End Sub

Sub DrawAtt(ByRef Surface As DirectDrawSurface4, attIndex As Byte, X As Byte, Y As Byte, XO As Long, YO As Long)
    TileSource.Left = ((attIndex - 1) Mod 7) * 32 + 8
    TileSource.Top = Int((attIndex - 1) / 7) * 32 + 8
    TileSource.Right = TileSource.Left + 16
    TileSource.Bottom = TileSource.Top + 16
    Call Surface.BltFast(X * 32 + XO, Y * 32 + YO, DDSAtts, TileSource, DDBLTFAST_SRCCOLORKEY)
End Sub

