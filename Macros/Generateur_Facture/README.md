# 📊 Macro_xlsx – Recherche Client & Génération de Facture

## Description

Ce mini projet permet de rechercher un client à partir de son ID et de **transférer automatiquement ses données** vers une feuille de facture. Cela facilite la **génération semi-automatique de factures** à partir d'une base client interne.

## ⚙️ Fonctionnalité principale

### 🔎 `demander_id()`

Cette macro demande à l'utilisateur d’entrer un **ID client**. Elle recherche cet ID dans une plage définie de la base de données clients. Si une correspondance est trouvée, elle copie les informations du client dans des cellules précises de la feuille **Facture**.

## 📋 Détail du fonctionnement

1. **Feuilles impliquées** :
   - `Facture` : Feuille de destination pour l’affichage des données.
   - `Bases de données clients` : Contient la base d’enregistrement des clients.

2. **Plage de recherche** :
   - Colonne **B5:B14** de la feuille *Bases de données clients* contenant les ids des commandes.

3. **Comportement utilisateur** :
   - Si aucun ID n’est saisi → alerte + arrêt de la macro.
   - Si l’ID est trouvé → les données sont transférées + message de confirmation.
   - Si l’ID n’est pas trouvé → message d’erreur.

## 🧪 Exemple de code VBA

```vba
Sub demander_id()

    Dim wsCommandes As Worksheet
    Dim wsFacture As Worksheet
    Dim myId As String
    Dim cell As Range
    Dim plageId As Range
    
    Set wsFacture = Sheets("Facture")
    Set wsCommandes = Sheets("Bases de données clients")
    
    ' Demander l'ID du client

    myId = Application.InputBox("Entrez l'ID du client à rechercher", "Recherche client")

    ' Vérifier si un id a été saisi

    If myId = "" Then
        MsgBox "ID non saisi.", vbExclamation
        Exit Sub
    End If
    
    ' Définir la plage de recherche dans la colonne B

    Set plageId = wsCommandes.Range("B5:B14")
    
    ' Chercher l'ID dans la plage

    For Each cell In plageId
        If cell.Value = myId Then

            ' Remplir la feuille Facture avec les données correspondantes

            wsFacture.Range("F4").Value = cell.Offset(0, 1).Value
            wsFacture.Range("F6").Value = cell.Offset(0, 2).Value
            wsFacture.Range("F8").Value = cell.Offset(0, 3).Value * cell.Offset(0, 4).Value
            MsgBox "Données client transférées avec succès.", vbInformation
            
            ' aller sur Facture

            wsFacture.Activate
            Exit Sub
        End If
    Next cell
    
    ' Si aucun ID trouvé

    MsgBox "ID introuvable dans la base de données.", vbCritical

End Sub
```
### Exporter la facture au format Pdf

```vba
Sub exportPDF()

 Sheets("Facture").ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:="facture.pdf", _
    Quality:=xlQualityStandard, _
    OpenAfterPublish:=True


End Sub
```

![image](https://github.com/user-attachments/assets/08280548-3f17-4f1f-ad57-97e7ce68a501)

![image](https://github.com/user-attachments/assets/d61069ae-eac7-49b8-bff8-13605c3798bf)

