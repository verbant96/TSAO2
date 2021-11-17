Attribute VB_Name = "modBalance"
Option Explicit

Private Type singleModification
    AFMago As Single
    AFGuerrero As Single
    AFCazador As Single
    AFPaladin As Single
    AFAsesino As Single
    AFBardo As Single
    AFClerigo As Single
    AFDruida As Single
    AMMago As Single
    AMGuerrero As Single
    AMCazador As Single
    AMPaladin As Single
    AMAsesino As Single
    AMBardo As Single
    AMClerigo As Single
    AMDruida As Single
End Type

Private Type singleClases
    Guerrero As singleModification
    Cazador As singleModification
    Paladin As singleModification
    Asesino As singleModification
    Ladron As singleModification
    Bardo As singleModification
    Clerigo As singleModification
    Mago As singleModification
    Druida As singleModification
End Type

Private Type balClases
    Guerrero As Single
    Cazador As Single
    Paladin As Single
    Asesino As Single
    Ladron As Single
    Bardo As Single
    Clerigo As Single
    Mago As Single
    Druida As Single
End Type

Private Type balRazaHP
    Humano(1 To 2) As Byte
    Elfo(1 To 2) As Byte
    ElfoOscuro(1 To 2) As Byte
    Gnomo(1 To 2) As Byte
    Enano(1 To 2) As Byte
End Type

Private Type balClassHP
    Guerrero As balRazaHP
    Cazador As balRazaHP
    Paladin As balRazaHP
    Asesino As balRazaHP
    Ladron As balRazaHP
    Bardo As balRazaHP
    Clerigo As balRazaHP
    Mago As balRazaHP
    Druida As balRazaHP
End Type

Private Type balModificaciones
    ModificadorEvasion As balClases
    ModificadorPoderAtaqueArmas As balClases
    ModificadorPoderAtaqueProyectiles As balClases
    ModicadorDañoClaseArmas As balClases
    ModicadorDañoClaseProyectiles As balClases
    ModEvasionDeEscudoClase As balClases
    AtaqueFisico As balClases
    AtaqueMagico As balClases
    DefensaFisica As balClases
    DefensaMagica As balClases
    AtaqueProyectil As balClases
    Vidas As balClassHP
End Type


Public Balance As balModificaciones
Private sBal As singleClases
Public Sub LoadBalance()

Dim l_file As clsIniReader
Dim d_file As clsIniReader

    Set l_file = New clsIniReader
    Set d_file = New clsIniReader

    '@ load file
    l_file.Initialize App.Path & "\Dat\Balance.dat"
    
    'Evasion escudo
    Balance.ModEvasionDeEscudoClase.Asesino = l_file.GetValue("ModEvasionDeEscudoClase", "Asesino")
    Balance.ModEvasionDeEscudoClase.Bardo = l_file.GetValue("ModEvasionDeEscudoClase", "Bardo")
    Balance.ModEvasionDeEscudoClase.Cazador = l_file.GetValue("ModEvasionDeEscudoClase", "Cazador")
    Balance.ModEvasionDeEscudoClase.Clerigo = l_file.GetValue("ModEvasionDeEscudoClase", "Clerigo")
    Balance.ModEvasionDeEscudoClase.Druida = l_file.GetValue("ModEvasionDeEscudoClase", "Druida")
    Balance.ModEvasionDeEscudoClase.Guerrero = l_file.GetValue("ModEvasionDeEscudoClase", "Guerrero")
    Balance.ModEvasionDeEscudoClase.Ladron = l_file.GetValue("ModEvasionDeEscudoClase", "Ladron")
    Balance.ModEvasionDeEscudoClase.Mago = l_file.GetValue("ModEvasionDeEscudoClase", "Mago")
    Balance.ModEvasionDeEscudoClase.Paladin = l_file.GetValue("ModEvasionDeEscudoClase", "Paladin")
    
    'Daño clases
    Balance.ModicadorDañoClaseArmas.Asesino = l_file.GetValue("ModicadorDañoClaseArmas", "Asesino")
    Balance.ModicadorDañoClaseArmas.Bardo = l_file.GetValue("ModicadorDañoClaseArmas", "Bardo")
    Balance.ModicadorDañoClaseArmas.Cazador = l_file.GetValue("ModicadorDañoClaseArmas", "Cazador")
    Balance.ModicadorDañoClaseArmas.Clerigo = l_file.GetValue("ModicadorDañoClaseArmas", "Clerigo")
    Balance.ModicadorDañoClaseArmas.Druida = l_file.GetValue("ModicadorDañoClaseArmas", "Druida")
    Balance.ModicadorDañoClaseArmas.Guerrero = l_file.GetValue("ModicadorDañoClaseArmas", "Guerrero")
    Balance.ModicadorDañoClaseArmas.Ladron = l_file.GetValue("ModicadorDañoClaseArmas", "Ladron")
    Balance.ModicadorDañoClaseArmas.Mago = l_file.GetValue("ModicadorDañoClaseArmas", "Mago")
    Balance.ModicadorDañoClaseArmas.Paladin = l_file.GetValue("ModicadorDañoClaseArmas", "Paladin")
    
    'Daño proyectiles
    Balance.ModicadorDañoClaseProyectiles.Asesino = l_file.GetValue("ModicadorDañoClaseProyectiles", "Asesino")
    Balance.ModicadorDañoClaseProyectiles.Bardo = l_file.GetValue("ModicadorDañoClaseProyectiles", "Bardo")
    Balance.ModicadorDañoClaseProyectiles.Cazador = l_file.GetValue("ModicadorDañoClaseProyectiles", "Cazador")
    Balance.ModicadorDañoClaseProyectiles.Clerigo = l_file.GetValue("ModicadorDañoClaseProyectiles", "Clerigo")
    Balance.ModicadorDañoClaseProyectiles.Druida = l_file.GetValue("ModicadorDañoClaseProyectiles", "Druida")
    Balance.ModicadorDañoClaseProyectiles.Guerrero = l_file.GetValue("ModicadorDañoClaseProyectiles", "Guerrero")
    Balance.ModicadorDañoClaseProyectiles.Ladron = l_file.GetValue("ModicadorDañoClaseProyectiles", "Ladron")
    Balance.ModicadorDañoClaseProyectiles.Mago = l_file.GetValue("ModicadorDañoClaseProyectiles", "Mago")
    Balance.ModicadorDañoClaseArmas.Paladin = l_file.GetValue("ModicadorDañoClaseProyectiles", "Paladin")
    
    'Evasion clase
    Balance.ModificadorEvasion.Asesino = l_file.GetValue("ModificadorEvasion", "Asesino")
    Balance.ModificadorEvasion.Bardo = l_file.GetValue("ModificadorEvasion", "Bardo")
    Balance.ModificadorEvasion.Cazador = l_file.GetValue("ModificadorEvasion", "Cazador")
    Balance.ModificadorEvasion.Clerigo = l_file.GetValue("ModificadorEvasion", "Clerigo")
    Balance.ModificadorEvasion.Druida = l_file.GetValue("ModificadorEvasion", "Druida")
    Balance.ModificadorEvasion.Guerrero = l_file.GetValue("ModificadorEvasion", "Guerrero")
    Balance.ModificadorEvasion.Ladron = l_file.GetValue("ModificadorEvasion", "Ladron")
    Balance.ModificadorEvasion.Mago = l_file.GetValue("ModificadorEvasion", "Mago")
    Balance.ModificadorEvasion.Paladin = l_file.GetValue("ModificadorEvasion", "Paladin")
    
    'Ataque c/armas
    Balance.ModificadorPoderAtaqueArmas.Asesino = l_file.GetValue("ModificadorPoderAtaqueArmas", "Asesino")
    Balance.ModificadorPoderAtaqueArmas.Bardo = l_file.GetValue("ModificadorPoderAtaqueArmas", "Bardo")
    Balance.ModificadorPoderAtaqueArmas.Cazador = l_file.GetValue("ModificadorPoderAtaqueArmas", "Cazador")
    Balance.ModificadorPoderAtaqueArmas.Clerigo = l_file.GetValue("ModificadorPoderAtaqueArmas", "Clerigo")
    Balance.ModificadorPoderAtaqueArmas.Druida = l_file.GetValue("ModificadorPoderAtaqueArmas", "Druida")
    Balance.ModificadorPoderAtaqueArmas.Guerrero = l_file.GetValue("ModificadorPoderAtaqueArmas", "Guerrero")
    Balance.ModificadorPoderAtaqueArmas.Ladron = l_file.GetValue("ModificadorPoderAtaqueArmas", "Ladron")
    Balance.ModificadorPoderAtaqueArmas.Mago = l_file.GetValue("ModificadorPoderAtaqueArmas", "Mago")
    Balance.ModificadorPoderAtaqueArmas.Paladin = l_file.GetValue("ModificadorPoderAtaqueArmas", "Paladin")
    
    'Ataque c/proyectiles
    Balance.ModificadorPoderAtaqueProyectiles.Asesino = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Asesino")
    Balance.ModificadorPoderAtaqueProyectiles.Bardo = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Bardo")
    Balance.ModificadorPoderAtaqueProyectiles.Cazador = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Cazador")
    Balance.ModificadorPoderAtaqueProyectiles.Clerigo = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Clerigo")
    Balance.ModificadorPoderAtaqueProyectiles.Druida = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Druida")
    Balance.ModificadorPoderAtaqueProyectiles.Guerrero = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Guerrero")
    Balance.ModificadorPoderAtaqueProyectiles.Ladron = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Ladron")
    Balance.ModificadorPoderAtaqueProyectiles.Mago = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Mago")
    Balance.ModificadorPoderAtaqueProyectiles.Paladin = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Paladin")
    
    'Ataque físico
    Balance.AtaqueFisico.Asesino = l_file.GetValue("AtaqueFisico", "Asesino")
    Balance.AtaqueFisico.Bardo = l_file.GetValue("AtaqueFisico", "Bardo")
    Balance.AtaqueFisico.Cazador = l_file.GetValue("AtaqueFisico", "Cazador")
    Balance.AtaqueFisico.Clerigo = l_file.GetValue("AtaqueFisico", "Clerigo")
    Balance.AtaqueFisico.Druida = l_file.GetValue("AtaqueFisico", "Druida")
    Balance.AtaqueFisico.Guerrero = l_file.GetValue("AtaqueFisico", "Guerrero")
    Balance.AtaqueFisico.Ladron = l_file.GetValue("AtaqueFisico", "Ladron")
    Balance.AtaqueFisico.Mago = l_file.GetValue("AtaqueFisico", "Mago")
    Balance.AtaqueFisico.Paladin = l_file.GetValue("AtaqueFisico", "Paladin")
    
    'Ataque mágico
    Balance.AtaqueMagico.Asesino = l_file.GetValue("AtaqueMagico", "Asesino")
    Balance.AtaqueMagico.Bardo = l_file.GetValue("AtaqueMagico", "Bardo")
    Balance.AtaqueMagico.Cazador = l_file.GetValue("AtaqueMagico", "Cazador")
    Balance.AtaqueMagico.Clerigo = l_file.GetValue("AtaqueMagico", "Clerigo")
    Balance.AtaqueMagico.Druida = l_file.GetValue("AtaqueMagico", "Druida")
    Balance.AtaqueMagico.Guerrero = l_file.GetValue("AtaqueMagico", "Guerrero")
    Balance.AtaqueMagico.Ladron = l_file.GetValue("AtaqueMagico", "Ladron")
    Balance.AtaqueMagico.Mago = l_file.GetValue("AtaqueMagico", "Mago")
    Balance.AtaqueMagico.Paladin = l_file.GetValue("AtaqueMagico", "Paladin")

    'Defensa fisica
    Balance.DefensaFisica.Asesino = l_file.GetValue("DefensaFisica", "Asesino")
    Balance.DefensaFisica.Bardo = l_file.GetValue("DefensaFisica", "Bardo")
    Balance.DefensaFisica.Cazador = l_file.GetValue("DefensaFisica", "Cazador")
    Balance.DefensaFisica.Clerigo = l_file.GetValue("DefensaFisica", "Clerigo")
    Balance.DefensaFisica.Druida = l_file.GetValue("DefensaFisica", "Druida")
    Balance.DefensaFisica.Guerrero = l_file.GetValue("DefensaFisica", "Guerrero")
    Balance.DefensaFisica.Ladron = l_file.GetValue("DefensaFisica", "Ladron")
    Balance.DefensaFisica.Mago = l_file.GetValue("DefensaFisica", "Mago")
    Balance.DefensaFisica.Paladin = l_file.GetValue("DefensaFisica", "Paladin")
    
    'Defensa mágica
    Balance.DefensaMagica.Asesino = l_file.GetValue("DefensaMagica", "Asesino")
    Balance.DefensaMagica.Bardo = l_file.GetValue("DefensaMagica", "Bardo")
    Balance.DefensaMagica.Cazador = l_file.GetValue("DefensaMagica", "Cazador")
    Balance.DefensaMagica.Clerigo = l_file.GetValue("DefensaMagica", "Clerigo")
    Balance.DefensaMagica.Druida = l_file.GetValue("DefensaMagica", "Druida")
    Balance.DefensaMagica.Guerrero = l_file.GetValue("DefensaMagica", "Guerrero")
    Balance.DefensaMagica.Ladron = l_file.GetValue("DefensaMagica", "Ladron")
    Balance.DefensaMagica.Mago = l_file.GetValue("DefensaMagica", "Mago")
    Balance.DefensaMagica.Paladin = l_file.GetValue("DefensaMagica", "Paladin")
    
    'Flechas
    Balance.AtaqueProyectil.Asesino = l_file.GetValue("AtaqueProyectil", "Asesino")
    Balance.AtaqueProyectil.Bardo = l_file.GetValue("AtaqueProyectil", "Bardo")
    Balance.AtaqueProyectil.Cazador = l_file.GetValue("AtaqueProyectil", "Cazador")
    Balance.AtaqueProyectil.Clerigo = l_file.GetValue("AtaqueProyectil", "Clerigo")
    Balance.AtaqueProyectil.Druida = l_file.GetValue("AtaqueProyectil", "Druida")
    Balance.AtaqueProyectil.Guerrero = l_file.GetValue("AtaqueProyectil", "Guerrero")
    Balance.AtaqueProyectil.Ladron = l_file.GetValue("AtaqueProyectil", "Ladron")
    Balance.AtaqueProyectil.Mago = l_file.GetValue("AtaqueProyectil", "Mago")
    Balance.AtaqueProyectil.Paladin = l_file.GetValue("AtaqueProyectil", "Paladin")
    
    
    'CLASE VS CLASE
    sBal.Asesino.AFAsesino = l_file.GetValue("Asesino", "AFAsesino")
    sBal.Asesino.AFBardo = l_file.GetValue("Asesino", "AFBardo")
    sBal.Asesino.AFCazador = l_file.GetValue("Asesino", "AFCazador")
    sBal.Asesino.AFClerigo = l_file.GetValue("Asesino", "AFClerigo")
    sBal.Asesino.AFDruida = l_file.GetValue("Asesino", "AFDruida")
    sBal.Asesino.AFGuerrero = l_file.GetValue("Asesino", "AFGuerrero")
    sBal.Asesino.AFMago = l_file.GetValue("Asesino", "AFMago")
    sBal.Asesino.AFPaladin = l_file.GetValue("Asesino", "AFPaladin")
    sBal.Asesino.AMAsesino = l_file.GetValue("Asesino", "AMAsesino")
    sBal.Asesino.AMBardo = l_file.GetValue("Asesino", "AMBardo")
    sBal.Asesino.AMCazador = l_file.GetValue("Asesino", "AMCazador")
    sBal.Asesino.AMClerigo = l_file.GetValue("Asesino", "AMClerigo")
    sBal.Asesino.AMDruida = l_file.GetValue("Asesino", "AMDruida")
    sBal.Asesino.AMGuerrero = l_file.GetValue("Asesino", "AMGuerrero")
    sBal.Asesino.AMMago = l_file.GetValue("Asesino", "AMMago")
    sBal.Asesino.AMPaladin = l_file.GetValue("Asesino", "AMPaladin")
    
    sBal.Bardo.AFAsesino = l_file.GetValue("Bardo", "AFAsesino")
    sBal.Bardo.AFBardo = l_file.GetValue("Bardo", "AFBardo")
    sBal.Bardo.AFCazador = l_file.GetValue("Bardo", "AFCazador")
    sBal.Bardo.AFClerigo = l_file.GetValue("Bardo", "AFClerigo")
    sBal.Bardo.AFDruida = l_file.GetValue("Bardo", "AFDruida")
    sBal.Bardo.AFGuerrero = l_file.GetValue("Bardo", "AFGuerrero")
    sBal.Bardo.AFMago = l_file.GetValue("Bardo", "AFMago")
    sBal.Bardo.AFPaladin = l_file.GetValue("Bardo", "AFPaladin")
    sBal.Bardo.AMAsesino = l_file.GetValue("Bardo", "AMAsesino")
    sBal.Bardo.AMBardo = l_file.GetValue("Bardo", "AMBardo")
    sBal.Bardo.AMCazador = l_file.GetValue("Bardo", "AMCazador")
    sBal.Bardo.AMClerigo = l_file.GetValue("Bardo", "AMClerigo")
    sBal.Bardo.AMDruida = l_file.GetValue("Bardo", "AMDruida")
    sBal.Bardo.AMGuerrero = l_file.GetValue("Bardo", "AMGuerrero")
    sBal.Bardo.AMMago = l_file.GetValue("Bardo", "AMMago")
    sBal.Bardo.AMPaladin = l_file.GetValue("Bardo", "AMPaladin")
    
    sBal.Cazador.AFAsesino = l_file.GetValue("Cazador", "AFAsesino")
    sBal.Cazador.AFBardo = l_file.GetValue("Cazador", "AFBardo")
    sBal.Cazador.AFCazador = l_file.GetValue("Cazador", "AFCazador")
    sBal.Cazador.AFClerigo = l_file.GetValue("Cazador", "AFClerigo")
    sBal.Cazador.AFDruida = l_file.GetValue("Cazador", "AFDruida")
    sBal.Cazador.AFGuerrero = l_file.GetValue("Cazador", "AFGuerrero")
    sBal.Cazador.AFMago = l_file.GetValue("Cazador", "AFMago")
    sBal.Cazador.AFPaladin = l_file.GetValue("Cazador", "AFPaladin")
    
    sBal.Clerigo.AFAsesino = l_file.GetValue("Clerigo", "AFAsesino")
    sBal.Clerigo.AFBardo = l_file.GetValue("Clerigo", "AFBardo")
    sBal.Clerigo.AFCazador = l_file.GetValue("Clerigo", "AFCazador")
    sBal.Clerigo.AFClerigo = l_file.GetValue("Clerigo", "AFClerigo")
    sBal.Clerigo.AFDruida = l_file.GetValue("Clerigo", "AFDruida")
    sBal.Clerigo.AFGuerrero = l_file.GetValue("Clerigo", "AFGuerrero")
    sBal.Clerigo.AFMago = l_file.GetValue("Clerigo", "AFMago")
    sBal.Clerigo.AFPaladin = l_file.GetValue("Clerigo", "AFPaladin")
    sBal.Clerigo.AMAsesino = l_file.GetValue("Clerigo", "AMAsesino")
    sBal.Clerigo.AMBardo = l_file.GetValue("Clerigo", "AMBardo")
    sBal.Clerigo.AMCazador = l_file.GetValue("Clerigo", "AMCazador")
    sBal.Clerigo.AMClerigo = l_file.GetValue("Clerigo", "AMClerigo")
    sBal.Clerigo.AMDruida = l_file.GetValue("Clerigo", "AMDruida")
    sBal.Clerigo.AMGuerrero = l_file.GetValue("Clerigo", "AMGuerrero")
    sBal.Clerigo.AMMago = l_file.GetValue("Clerigo", "AMMago")
    sBal.Clerigo.AMPaladin = l_file.GetValue("Clerigo", "AMPaladin")
    
    sBal.Druida.AFAsesino = l_file.GetValue("Druida", "AFAsesino")
    sBal.Druida.AFBardo = l_file.GetValue("Druida", "AFBardo")
    sBal.Druida.AFCazador = l_file.GetValue("Druida", "AFCazador")
    sBal.Druida.AFClerigo = l_file.GetValue("Druida", "AFClerigo")
    sBal.Druida.AFDruida = l_file.GetValue("Druida", "AFDruida")
    sBal.Druida.AFGuerrero = l_file.GetValue("Druida", "AFGuerrero")
    sBal.Druida.AFMago = l_file.GetValue("Druida", "AFMago")
    sBal.Druida.AFPaladin = l_file.GetValue("Druida", "AFPaladin")
    sBal.Druida.AMAsesino = l_file.GetValue("Druida", "AMAsesino")
    sBal.Druida.AMBardo = l_file.GetValue("Druida", "AMBardo")
    sBal.Druida.AMCazador = l_file.GetValue("Druida", "AMCazador")
    sBal.Druida.AMClerigo = l_file.GetValue("Druida", "AMClerigo")
    sBal.Druida.AMDruida = l_file.GetValue("Druida", "AMDruida")
    sBal.Druida.AMGuerrero = l_file.GetValue("Druida", "AMGuerrero")
    sBal.Druida.AMMago = l_file.GetValue("Druida", "AMMago")
    sBal.Druida.AMPaladin = l_file.GetValue("Druida", "AMPaladin")
    
    sBal.Guerrero.AFAsesino = l_file.GetValue("Guerrero", "AFAsesino")
    sBal.Guerrero.AFBardo = l_file.GetValue("Guerrero", "AFBardo")
    sBal.Guerrero.AFCazador = l_file.GetValue("Guerrero", "AFCazador")
    sBal.Guerrero.AFClerigo = l_file.GetValue("Guerrero", "AFClerigo")
    sBal.Guerrero.AFDruida = l_file.GetValue("Guerrero", "AFDruida")
    sBal.Guerrero.AFGuerrero = l_file.GetValue("Guerrero", "AFGuerrero")
    sBal.Guerrero.AFMago = l_file.GetValue("Guerrero", "AFMago")
    sBal.Guerrero.AFPaladin = l_file.GetValue("Guerrero", "AFPaladin")

    sBal.Mago.AMAsesino = l_file.GetValue("Mago", "AMAsesino")
    sBal.Mago.AMBardo = l_file.GetValue("Mago", "AMBardo")
    sBal.Mago.AMCazador = l_file.GetValue("Mago", "AMCazador")
    sBal.Mago.AMClerigo = l_file.GetValue("Mago", "AMClerigo")
    sBal.Mago.AMDruida = l_file.GetValue("Mago", "AMDruida")
    sBal.Mago.AMGuerrero = l_file.GetValue("Mago", "AMGuerrero")
    sBal.Mago.AMMago = l_file.GetValue("Mago", "AMMago")
    sBal.Mago.AMPaladin = l_file.GetValue("Mago", "AMPaladin")
    
    sBal.Paladin.AFAsesino = l_file.GetValue("Paladin", "AFAsesino")
    sBal.Paladin.AFBardo = l_file.GetValue("Paladin", "AFBardo")
    sBal.Paladin.AFCazador = l_file.GetValue("Paladin", "AFCazador")
    sBal.Paladin.AFClerigo = l_file.GetValue("Paladin", "AFClerigo")
    sBal.Paladin.AFDruida = l_file.GetValue("Paladin", "AFDruida")
    sBal.Paladin.AFGuerrero = l_file.GetValue("Paladin", "AFGuerrero")
    sBal.Paladin.AFMago = l_file.GetValue("Paladin", "AFMago")
    sBal.Paladin.AFPaladin = l_file.GetValue("Paladin", "AFPaladin")
    sBal.Paladin.AMAsesino = l_file.GetValue("Paladin", "AMAsesino")
    sBal.Paladin.AMBardo = l_file.GetValue("Paladin", "AMBardo")
    sBal.Paladin.AMCazador = l_file.GetValue("Paladin", "AMCazador")
    sBal.Paladin.AMClerigo = l_file.GetValue("Paladin", "AMClerigo")
    sBal.Paladin.AMDruida = l_file.GetValue("Paladin", "AMDruida")
    sBal.Paladin.AMGuerrero = l_file.GetValue("Paladin", "AMGuerrero")
    sBal.Paladin.AMMago = l_file.GetValue("Paladin", "AMMago")
    sBal.Paladin.AMPaladin = l_file.GetValue("Paladin", "AMPaladin")
    
    
    'Vidas
    d_file.Initialize App.Path & "\Dat\Vidas.dat"
    Dim tmpV As String
        
        'VIDAS MAGO
        tmpV = d_file.GetValue("MAGO", "Humano")
        Balance.Vidas.Mago.Humano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Mago.Humano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("MAGO", "Elfo")
        Balance.Vidas.Mago.Elfo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Mago.Elfo(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("MAGO", "ElfoOscuro")
        Balance.Vidas.Mago.ElfoOscuro(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Mago.ElfoOscuro(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("MAGO", "Enano")
        Balance.Vidas.Mago.Enano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Mago.Enano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("MAGO", "Gnomo")
        Balance.Vidas.Mago.Gnomo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Mago.Gnomo(2) = ReadField(2, tmpV, Asc("-"))
        
        
        'VIDAS CLERIGO
        tmpV = d_file.GetValue("CLERIGO", "Humano")
        Balance.Vidas.Clerigo.Humano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Clerigo.Humano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("CLERIGO", "Elfo")
        Balance.Vidas.Clerigo.Elfo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Clerigo.Elfo(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("CLERIGO", "ElfoOscuro")
        Balance.Vidas.Clerigo.ElfoOscuro(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Clerigo.ElfoOscuro(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("CLERIGO", "Enano")
        Balance.Vidas.Clerigo.Enano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Clerigo.Enano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("CLERIGO", "Gnomo")
        Balance.Vidas.Clerigo.Gnomo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Clerigo.Gnomo(2) = ReadField(2, tmpV, Asc("-"))
        
        
        'VIDAS BARDO
        tmpV = d_file.GetValue("BARDO", "Humano")
        Balance.Vidas.Bardo.Humano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Bardo.Humano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("BARDO", "Elfo")
        Balance.Vidas.Bardo.Elfo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Bardo.Elfo(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("BARDO", "ElfoOscuro")
        Balance.Vidas.Bardo.ElfoOscuro(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Bardo.ElfoOscuro(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("BARDO", "Enano")
        Balance.Vidas.Bardo.Enano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Bardo.Enano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("BARDO", "Gnomo")
        Balance.Vidas.Bardo.Gnomo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Bardo.Gnomo(2) = ReadField(2, tmpV, Asc("-"))
    
    
    
        'VIDAS DRUIDA
        tmpV = d_file.GetValue("DRUIDA", "Humano")
        Balance.Vidas.Druida.Humano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Druida.Humano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("DRUIDA", "Elfo")
        Balance.Vidas.Druida.Elfo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Druida.Elfo(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("DRUIDA", "ElfoOscuro")
        Balance.Vidas.Druida.ElfoOscuro(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Druida.ElfoOscuro(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("DRUIDA", "Enano")
        Balance.Vidas.Druida.Enano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Druida.Enano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("DRUIDA", "Gnomo")
        Balance.Vidas.Druida.Gnomo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Druida.Gnomo(2) = ReadField(2, tmpV, Asc("-"))
        
        
        'VIDAS PALADIN
        tmpV = d_file.GetValue("PALADIN", "Humano")
        Balance.Vidas.Paladin.Humano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Paladin.Humano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("PALADIN", "Elfo")
        Balance.Vidas.Paladin.Elfo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Paladin.Elfo(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("PALADIN", "ElfoOscuro")
        Balance.Vidas.Paladin.ElfoOscuro(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Paladin.ElfoOscuro(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("PALADIN", "Enano")
        Balance.Vidas.Paladin.Enano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Paladin.Enano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("PALADIN", "Gnomo")
        Balance.Vidas.Paladin.Gnomo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Paladin.Gnomo(2) = ReadField(2, tmpV, Asc("-"))
        
        
        'VIDAS ASESINO
        tmpV = d_file.GetValue("ASESINO", "Humano")
        Balance.Vidas.Asesino.Humano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Asesino.Humano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("ASESINO", "Elfo")
        Balance.Vidas.Asesino.Elfo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Asesino.Elfo(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("ASESINO", "ElfoOscuro")
        Balance.Vidas.Asesino.ElfoOscuro(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Asesino.ElfoOscuro(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("ASESINO", "Enano")
        Balance.Vidas.Asesino.Enano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Asesino.Enano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("ASESINO", "Gnomo")
        Balance.Vidas.Asesino.Gnomo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Asesino.Gnomo(2) = ReadField(2, tmpV, Asc("-"))
        
        
        
        'VIDAS GUERRERO
        tmpV = d_file.GetValue("GUERRERO", "Humano")
        Balance.Vidas.Guerrero.Humano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Guerrero.Humano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("GUERRERO", "Elfo")
        Balance.Vidas.Guerrero.Elfo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Guerrero.Elfo(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("GUERRERO", "ElfoOscuro")
        Balance.Vidas.Guerrero.ElfoOscuro(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Guerrero.ElfoOscuro(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("GUERRERO", "Enano")
        Balance.Vidas.Guerrero.Enano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Guerrero.Enano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("GUERRERO", "Gnomo")
        Balance.Vidas.Guerrero.Gnomo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Guerrero.Gnomo(2) = ReadField(2, tmpV, Asc("-"))
        
        
        
        'VIDAS CAZADOR
        tmpV = d_file.GetValue("CAZADOR", "Humano")
        Balance.Vidas.Cazador.Humano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Cazador.Humano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("CAZADOR", "Elfo")
        Balance.Vidas.Cazador.Elfo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Cazador.Elfo(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("CAZADOR", "ElfoOscuro")
        Balance.Vidas.Cazador.ElfoOscuro(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Cazador.ElfoOscuro(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("CAZADOR", "Enano")
        Balance.Vidas.Cazador.Enano(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Cazador.Enano(2) = ReadField(2, tmpV, Asc("-"))
        
        tmpV = d_file.GetValue("CAZADOR", "Gnomo")
        Balance.Vidas.Cazador.Gnomo(1) = ReadField(1, tmpV, Asc("-"))
        Balance.Vidas.Cazador.Gnomo(2) = ReadField(2, tmpV, Asc("-"))
    
    
End Sub
Function ModificarAtaqueProyectil(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Guerrero
        Case "CAZADOR"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Cazador
        Case "PALADIN"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Paladin
        Case "ASESINO"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Asesino
        Case "LADRON"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Ladron
        Case "BARDO"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Bardo
        Case "CLERIGO"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Clerigo
        Case "MAGO"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Mago
        Case "DRUIDA"
            ModificarAtaqueProyectil = Balance.AtaqueProyectil.Druida
        Case Else
            ModificarAtaqueProyectil = 0
    End Select
    
End Function
Function ModificarAtaqueFisico(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Guerrero
        Case "CAZADOR"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Cazador
        Case "PALADIN"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Paladin
        Case "ASESINO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Asesino
        Case "LADRON"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Ladron
        Case "BARDO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Bardo
        Case "CLERIGO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Clerigo
        Case "MAGO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Mago
        Case "DRUIDA"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Druida
        Case Else
            ModificarAtaqueFisico = 0
    End Select
    
End Function
Function ModificarAtaqueMagico(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Guerrero
        Case "CAZADOR"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Cazador
        Case "PALADIN"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Paladin
        Case "ASESINO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Asesino
        Case "LADRON"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Ladron
        Case "BARDO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Bardo
        Case "CLERIGO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Clerigo
        Case "MAGO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Mago
        Case "DRUIDA"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Druida
        Case Else
            ModificarAtaqueMagico = 0
    End Select
    
End Function
Function ModificarDefensaFisica(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarDefensaFisica = Balance.DefensaFisica.Guerrero
        Case "CAZADOR"
            ModificarDefensaFisica = Balance.DefensaFisica.Cazador
        Case "PALADIN"
            ModificarDefensaFisica = Balance.DefensaFisica.Paladin
        Case "ASESINO"
            ModificarDefensaFisica = Balance.DefensaFisica.Asesino
        Case "LADRON"
            ModificarDefensaFisica = Balance.DefensaFisica.Ladron
        Case "BARDO"
            ModificarDefensaFisica = Balance.DefensaFisica.Bardo
        Case "CLERIGO"
            ModificarDefensaFisica = Balance.DefensaFisica.Clerigo
        Case "MAGO"
            ModificarDefensaFisica = Balance.DefensaFisica.Mago
        Case "DRUIDA"
            ModificarDefensaFisica = Balance.DefensaFisica.Druida
        Case Else
            ModificarDefensaFisica = 0
    End Select
    
End Function
Function ModificarDefensaMagica(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarDefensaMagica = Balance.DefensaMagica.Guerrero
        Case "CAZADOR"
            ModificarDefensaMagica = Balance.DefensaMagica.Cazador
        Case "PALADIN"
            ModificarDefensaMagica = Balance.DefensaMagica.Paladin
        Case "ASESINO"
            ModificarDefensaMagica = Balance.DefensaMagica.Asesino
        Case "LADRON"
            ModificarDefensaMagica = Balance.DefensaMagica.Ladron
        Case "BARDO"
            ModificarDefensaMagica = Balance.DefensaMagica.Bardo
        Case "CLERIGO"
            ModificarDefensaMagica = Balance.DefensaMagica.Clerigo
        Case "MAGO"
            ModificarDefensaMagica = Balance.DefensaMagica.Mago
        Case "DRUIDA"
            ModificarDefensaMagica = Balance.DefensaMagica.Druida
        Case Else
            ModificarDefensaMagica = 0
    End Select
    
End Function
Function ModificarAFClasevsClase(ByVal cAtacante As String, ByVal cVictima As String) As Single

    Select Case UCase$(cAtacante)
        Case "GUERRERO"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAFClasevsClase = sBal.Guerrero.AFGuerrero
                Case "CAZADOR"
                    ModificarAFClasevsClase = sBal.Guerrero.AFCazador
                Case "PALADIN"
                    ModificarAFClasevsClase = sBal.Guerrero.AFPaladin
                Case "ASESINO"
                    ModificarAFClasevsClase = sBal.Guerrero.AFAsesino
                Case "BARDO"
                    ModificarAFClasevsClase = sBal.Guerrero.AFBardo
                Case "CLERIGO"
                    ModificarAFClasevsClase = sBal.Guerrero.AFClerigo
                Case "MAGO"
                    ModificarAFClasevsClase = sBal.Guerrero.AFMago
                Case "DRUIDA"
                    ModificarAFClasevsClase = sBal.Guerrero.AFDruida
                Case Else
                    ModificarAFClasevsClase = 0
            End Select
            
            
        Case "CAZADOR"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAFClasevsClase = sBal.Cazador.AFGuerrero
                Case "CAZADOR"
                    ModificarAFClasevsClase = sBal.Cazador.AFCazador
                Case "PALADIN"
                    ModificarAFClasevsClase = sBal.Cazador.AFPaladin
                Case "ASESINO"
                    ModificarAFClasevsClase = sBal.Cazador.AFAsesino
                Case "BARDO"
                    ModificarAFClasevsClase = sBal.Cazador.AFBardo
                Case "CLERIGO"
                    ModificarAFClasevsClase = sBal.Cazador.AFClerigo
                Case "MAGO"
                    ModificarAFClasevsClase = sBal.Cazador.AFMago
                Case "DRUIDA"
                    ModificarAFClasevsClase = sBal.Cazador.AFDruida
                Case Else
                    ModificarAFClasevsClase = 0
            End Select
            
            
        Case "PALADIN"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAFClasevsClase = sBal.Paladin.AFGuerrero
                Case "CAZADOR"
                    ModificarAFClasevsClase = sBal.Paladin.AFCazador
                Case "PALADIN"
                    ModificarAFClasevsClase = sBal.Paladin.AFPaladin
                Case "ASESINO"
                    ModificarAFClasevsClase = sBal.Paladin.AFAsesino
                Case "BARDO"
                    ModificarAFClasevsClase = sBal.Paladin.AFBardo
                Case "CLERIGO"
                    ModificarAFClasevsClase = sBal.Paladin.AFClerigo
                Case "MAGO"
                    ModificarAFClasevsClase = sBal.Paladin.AFMago
                Case "DRUIDA"
                    ModificarAFClasevsClase = sBal.Paladin.AFDruida
                Case Else
                    ModificarAFClasevsClase = 0
            End Select
            
            
        Case "ASESINO"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAFClasevsClase = sBal.Asesino.AFGuerrero
                Case "CAZADOR"
                    ModificarAFClasevsClase = sBal.Asesino.AFCazador
                Case "PALADIN"
                    ModificarAFClasevsClase = sBal.Asesino.AFPaladin
                Case "ASESINO"
                    ModificarAFClasevsClase = sBal.Asesino.AFAsesino
                Case "BARDO"
                    ModificarAFClasevsClase = sBal.Asesino.AFBardo
                Case "CLERIGO"
                    ModificarAFClasevsClase = sBal.Asesino.AFClerigo
                Case "MAGO"
                    ModificarAFClasevsClase = sBal.Asesino.AFMago
                Case "DRUIDA"
                    ModificarAFClasevsClase = sBal.Asesino.AFDruida
                Case Else
                    ModificarAFClasevsClase = 0
            End Select
            
            
        Case "BARDO"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAFClasevsClase = sBal.Bardo.AFGuerrero
                Case "CAZADOR"
                    ModificarAFClasevsClase = sBal.Bardo.AFCazador
                Case "PALADIN"
                    ModificarAFClasevsClase = sBal.Bardo.AFPaladin
                Case "ASESINO"
                    ModificarAFClasevsClase = sBal.Bardo.AFAsesino
                Case "BARDO"
                    ModificarAFClasevsClase = sBal.Bardo.AFBardo
                Case "CLERIGO"
                    ModificarAFClasevsClase = sBal.Bardo.AFClerigo
                Case "MAGO"
                    ModificarAFClasevsClase = sBal.Bardo.AFMago
                Case "DRUIDA"
                    ModificarAFClasevsClase = sBal.Bardo.AFDruida
                Case Else
                    ModificarAFClasevsClase = 0
            End Select
            
            
        Case "CLERIGO"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAFClasevsClase = sBal.Clerigo.AFGuerrero
                Case "CAZADOR"
                    ModificarAFClasevsClase = sBal.Clerigo.AFCazador
                Case "PALADIN"
                    ModificarAFClasevsClase = sBal.Clerigo.AFPaladin
                Case "ASESINO"
                    ModificarAFClasevsClase = sBal.Clerigo.AFAsesino
                Case "BARDO"
                    ModificarAFClasevsClase = sBal.Clerigo.AFBardo
                Case "CLERIGO"
                    ModificarAFClasevsClase = sBal.Clerigo.AFClerigo
                Case "MAGO"
                    ModificarAFClasevsClase = sBal.Clerigo.AFMago
                Case "DRUIDA"
                    ModificarAFClasevsClase = sBal.Clerigo.AFDruida
                Case Else
                    ModificarAFClasevsClase = 0
            End Select
            
            
        Case "DRUIDA"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAFClasevsClase = sBal.Druida.AFGuerrero
                Case "CAZADOR"
                    ModificarAFClasevsClase = sBal.Druida.AFCazador
                Case "PALADIN"
                    ModificarAFClasevsClase = sBal.Druida.AFPaladin
                Case "ASESINO"
                    ModificarAFClasevsClase = sBal.Druida.AFAsesino
                Case "BARDO"
                    ModificarAFClasevsClase = sBal.Druida.AFBardo
                Case "CLERIGO"
                    ModificarAFClasevsClase = sBal.Druida.AFClerigo
                Case "MAGO"
                    ModificarAFClasevsClase = sBal.Druida.AFMago
                Case "DRUIDA"
                    ModificarAFClasevsClase = sBal.Druida.AFDruida
                Case Else
                    ModificarAFClasevsClase = 0
            End Select
            
            
        Case Else
            ModificarAFClasevsClase = 0
    End Select

End Function
Function ModificarAMClasevsClase(ByVal cAtacante As String, ByVal cVictima As String) As Single

    Select Case UCase$(cAtacante)
        Case "MAGO"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAMClasevsClase = sBal.Mago.AMGuerrero
                Case "CAZADOR"
                    ModificarAMClasevsClase = sBal.Mago.AMCazador
                Case "PALADIN"
                    ModificarAMClasevsClase = sBal.Mago.AMPaladin
                Case "ASESINO"
                    ModificarAMClasevsClase = sBal.Mago.AMAsesino
                Case "BARDO"
                    ModificarAMClasevsClase = sBal.Mago.AMBardo
                Case "CLERIGO"
                    ModificarAMClasevsClase = sBal.Mago.AMClerigo
                Case "MAGO"
                    ModificarAMClasevsClase = sBal.Mago.AMMago
                Case "DRUIDA"
                    ModificarAMClasevsClase = sBal.Mago.AMDruida
                Case Else
                    ModificarAMClasevsClase = 0
            End Select
            
            
        Case "PALADIN"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAMClasevsClase = sBal.Paladin.AMGuerrero
                Case "CAZADOR"
                    ModificarAMClasevsClase = sBal.Paladin.AMCazador
                Case "PALADIN"
                    ModificarAMClasevsClase = sBal.Paladin.AMPaladin
                Case "ASESINO"
                    ModificarAMClasevsClase = sBal.Paladin.AMAsesino
                Case "BARDO"
                    ModificarAMClasevsClase = sBal.Paladin.AMBardo
                Case "CLERIGO"
                    ModificarAMClasevsClase = sBal.Paladin.AMClerigo
                Case "MAGO"
                    ModificarAMClasevsClase = sBal.Paladin.AMMago
                Case "DRUIDA"
                    ModificarAMClasevsClase = sBal.Paladin.AMDruida
                Case Else
                    ModificarAMClasevsClase = 0
            End Select
            
            
        Case "ASESINO"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAMClasevsClase = sBal.Asesino.AMGuerrero
                Case "CAZADOR"
                    ModificarAMClasevsClase = sBal.Asesino.AMCazador
                Case "PALADIN"
                    ModificarAMClasevsClase = sBal.Asesino.AMPaladin
                Case "ASESINO"
                    ModificarAMClasevsClase = sBal.Asesino.AMAsesino
                Case "BARDO"
                    ModificarAMClasevsClase = sBal.Asesino.AMBardo
                Case "CLERIGO"
                    ModificarAMClasevsClase = sBal.Asesino.AMClerigo
                Case "MAGO"
                    ModificarAMClasevsClase = sBal.Asesino.AMMago
                Case "DRUIDA"
                    ModificarAMClasevsClase = sBal.Asesino.AMDruida
                Case Else
                    ModificarAMClasevsClase = 0
            End Select
            
            
        Case "BARDO"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAMClasevsClase = sBal.Bardo.AMGuerrero
                Case "CAZADOR"
                    ModificarAMClasevsClase = sBal.Bardo.AMCazador
                Case "PALADIN"
                    ModificarAMClasevsClase = sBal.Bardo.AMPaladin
                Case "ASESINO"
                    ModificarAMClasevsClase = sBal.Bardo.AMAsesino
                Case "BARDO"
                    ModificarAMClasevsClase = sBal.Bardo.AMBardo
                Case "CLERIGO"
                    ModificarAMClasevsClase = sBal.Bardo.AMClerigo
                Case "MAGO"
                    ModificarAMClasevsClase = sBal.Bardo.AMMago
                Case "DRUIDA"
                    ModificarAMClasevsClase = sBal.Bardo.AMDruida
                Case Else
                    ModificarAMClasevsClase = 0
            End Select
            
            
        Case "CLERIGO"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAMClasevsClase = sBal.Clerigo.AMGuerrero
                Case "CAZADOR"
                    ModificarAMClasevsClase = sBal.Clerigo.AMCazador
                Case "PALADIN"
                    ModificarAMClasevsClase = sBal.Clerigo.AMPaladin
                Case "ASESINO"
                    ModificarAMClasevsClase = sBal.Clerigo.AMAsesino
                Case "BARDO"
                    ModificarAMClasevsClase = sBal.Clerigo.AMBardo
                Case "CLERIGO"
                    ModificarAMClasevsClase = sBal.Clerigo.AMClerigo
                Case "MAGO"
                    ModificarAMClasevsClase = sBal.Clerigo.AMMago
                Case "DRUIDA"
                    ModificarAMClasevsClase = sBal.Clerigo.AMDruida
                Case Else
                    ModificarAMClasevsClase = 0
            End Select
            
            
        Case "DRUIDA"
            Select Case UCase$(cVictima)
                Case "GUERRERO"
                    ModificarAMClasevsClase = sBal.Druida.AMGuerrero
                Case "CAZADOR"
                    ModificarAMClasevsClase = sBal.Druida.AMCazador
                Case "PALADIN"
                    ModificarAMClasevsClase = sBal.Druida.AMPaladin
                Case "ASESINO"
                    ModificarAMClasevsClase = sBal.Druida.AMAsesino
                Case "BARDO"
                    ModificarAMClasevsClase = sBal.Druida.AMBardo
                Case "CLERIGO"
                    ModificarAMClasevsClase = sBal.Druida.AMClerigo
                Case "MAGO"
                    ModificarAMClasevsClase = sBal.Druida.AMMago
                Case "DRUIDA"
                    ModificarAMClasevsClase = sBal.Druida.AMDruida
                Case Else
                    ModificarAMClasevsClase = 0
            End Select
            
            
        Case Else
            ModificarAMClasevsClase = 0
    End Select

End Function
