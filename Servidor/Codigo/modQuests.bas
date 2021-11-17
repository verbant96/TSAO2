Attribute VB_Name = "modQuests"
Option Explicit
Public QuestsList() As tQuests
Public Type tQuests
    Name As String
    Tipo As Byte
    Usuarios As Integer
    ptsTorneo As Integer
    ptsTS As Integer
    Creditos As Integer
    Oro As Long
    numNPC As Integer
    CantNPC As Integer
End Type
Public Sub CargarQuests()
        Dim p As Integer, loopC As Integer, LoopD
        p = val(GetVar(App.Path & "\Dat\QUESTS.dat", "INIT", "Num"))
   
        ReDim QuestsList(p) As tQuests
       
        For loopC = 1 To p
            QuestsList(loopC).Name = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "Name")
            QuestsList(loopC).Tipo = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "Tipo")
            QuestsList(loopC).ptsTorneo = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "ptsTorneo")
            QuestsList(loopC).Creditos = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "Creditos")
            QuestsList(loopC).ptsTS = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "ptsTS")
            QuestsList(loopC).Oro = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "Oro")
            
            If QuestsList(loopC).Tipo = 1 Then
                QuestsList(loopC).numNPC = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "MataNPC")
                QuestsList(loopC).CantNPC = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "Cant")
            ElseIf QuestsList(loopC).Tipo = 2 Then
                QuestsList(loopC).Usuarios = GetVar(App.Path & "\Dat\QUESTS.dat", "Quest" & loopC, "Usuarios")
            End If
                
        Next loopC
End Sub
Public Sub RestarNPC(ByVal userindex As Integer, ByVal KillNPC As Integer)

    Dim NroQuest As Byte, i As Long, CompletoQuest As Boolean
    NroQuest = UserList(userindex).flags.UserNumQuest

    If QuestsList(NroQuest).Tipo = 1 Then
    
            If KillNPC = QuestsList(NroQuest).numNPC Then
                UserList(userindex).flags.MuereQuest = UserList(userindex).flags.MuereQuest + 1
            End If
            
            'Si la completo, gg
            If UserList(userindex).flags.MuereQuest >= QuestsList(NroQuest).CantNPC Then
                CompletoQuest = True
            Else
                CompletoQuest = False
            End If
            
       If CompletoQuest = True Then
       
            Call SendData(SendTarget.toindex, userindex, 0, "||66")
            
            Dim tmpReward As Long
            
            'Oro
            If QuestsList(NroQuest).Oro > 0 Then
                tmpReward = IIf((UserList(userindex).flags.estado = 0 And UserList(userindex).flags.EsPremium = 0), val(QuestsList(NroQuest).Oro), val(QuestsList(NroQuest).Oro) * 2)
                UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + tmpReward
                Call SendData(SendTarget.toindex, userindex, 0, "||63@" & PonerPuntos(tmpReward))
            End If
            
            'Puntos de Torneo
            If QuestsList(NroQuest).ptsTorneo > 0 Then
                tmpReward = IIf((UserList(userindex).flags.estado = 0 And UserList(userindex).flags.EsPremium = 0), val(QuestsList(NroQuest).ptsTorneo), val(QuestsList(NroQuest).ptsTorneo) * 2)
                Call SendData(SendTarget.toindex, userindex, 0, "||57@" & tmpReward)
                Call AgregarPuntos(userindex, (val(tmpReward)))
            End If
            
            'TS Points
            If QuestsList(NroQuest).ptsTS > 0 Then
                tmpReward = val(QuestsList(NroQuest).ptsTS)
                UserList(userindex).Stats.TSPoints = UserList(userindex).Stats.TSPoints + (tmpReward)
                Call SendData(SendTarget.toindex, userindex, 0, "||900@" & tmpReward)
            End If
            
            'Creditos
            If QuestsList(NroQuest).Creditos > 0 Then
                tmpReward = val(QuestsList(NroQuest).Creditos)
                UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion + tmpReward
                Call SendData(SendTarget.toindex, userindex, 0, "||930@" & tmpReward)
            End If
            
            'Reputación
            tmpReward = IIf((UserList(userindex).flags.estado = 0 And UserList(userindex).flags.EsPremium = 0), val(QuestsList(NroQuest).ptsTorneo), val(QuestsList(NroQuest).ptsTorneo) * 2)
            UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + tmpReward
                    
            modQuests.ResetQuest (userindex)
            UserList(userindex).flags.QuestCompletadas = UserList(userindex).flags.QuestCompletadas + 1
            
            SendUserGLD (userindex)
    End If
  End If

End Sub
Public Sub RestarUser(ByVal userindex As Integer, ByVal VictimIndex As Integer)

    Dim NroQuest As Byte
    NroQuest = UserList(userindex).flags.UserNumQuest

    If QuestsList(NroQuest).Tipo = 2 Then
        If UserList(userindex).flags.Questeando = 1 And TriggerZonaPelea(userindex, VictimIndex) <> TRIGGER6_PERMITE Then
            UserList(userindex).flags.MuereQuest = UserList(userindex).flags.MuereQuest + 1
        End If
         
        If UserList(userindex).flags.MuereQuest = QuestsList(NroQuest).Usuarios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||66")
            
            Dim tmpReward As Long
            
            'Oro
            If QuestsList(NroQuest).Oro > 0 Then
                tmpReward = IIf(UserList(userindex).flags.estado, val(QuestsList(NroQuest).Oro), val(QuestsList(NroQuest).Oro) * 2)
                UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + tmpReward
                Call SendData(SendTarget.toindex, userindex, 0, "||63@" & PonerPuntos(tmpReward))
            End If
            
            'Puntos de Torneo
            If QuestsList(NroQuest).ptsTorneo > 0 Then
                tmpReward = IIf(UserList(userindex).flags.estado, val(QuestsList(NroQuest).ptsTorneo), val(QuestsList(NroQuest).ptsTorneo) * 2)
                Call SendData(SendTarget.toindex, userindex, 0, "||57@" & tmpReward)
                Call AgregarPuntos(userindex, (val(tmpReward)))
            End If
            
            'TS Points
            If QuestsList(NroQuest).ptsTS > 0 Then
                tmpReward = IIf(UserList(userindex).flags.estado, val(QuestsList(NroQuest).ptsTS), val(QuestsList(NroQuest).ptsTS) * 2)
                UserList(userindex).Stats.TSPoints = UserList(userindex).Stats.TSPoints + (tmpReward)
                Call SendData(SendTarget.toindex, userindex, 0, "||900@" & tmpReward)
            End If
            
            'Creditos
            If QuestsList(NroQuest).Creditos > 0 Then
                tmpReward = IIf(UserList(userindex).flags.estado, val(QuestsList(NroQuest).Creditos), val(QuestsList(NroQuest).Creditos) * 2)
                UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion + tmpReward
                Call SendData(SendTarget.toindex, userindex, 0, "||930@" & tmpReward)
            End If
            
            'Reputación
            tmpReward = IIf(UserList(userindex).flags.estado, val(QuestsList(NroQuest).ptsTorneo), val(QuestsList(NroQuest).ptsTorneo) * 2)
            UserList(userindex).Stats.Reputacione = UserList(userindex).Stats.Reputacione + tmpReward
                    
            modQuests.ResetQuest (userindex)
            UserList(userindex).flags.QuestCompletadas = UserList(userindex).flags.QuestCompletadas + 1
            
            SendUserGLD (userindex)
        End If
    End If


End Sub
Public Sub ResetQuest(ByVal userindex As Integer)

        UserList(userindex).flags.MuereQuest = 0
        UserList(userindex).flags.Questeando = 0
        UserList(userindex).flags.UserNumQuest = 0
        
End Sub
Public Function SendQuestList(ByVal userindex As Integer) As String
Dim tStr As String, tIntx As Integer
 
    tStr = UBound(QuestsList) & ","
    For tIntx = 1 To UBound(QuestsList)
        tStr = tStr & QuestsList(tIntx).Name & ","
    Next tIntx
    
    SendQuestList = tStr
End Function
