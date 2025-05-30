# 📋 Mini gestionnaire de tâches – Macro_xlsx

## Description

**Mini gestionnaire de tâches** est un projet Excel avec macros VBA permettant de gérer des tâches simplement depuis une interface de formulaire. Il permet pour l’instant d’ajouter une tâche à une liste, avec des champs tels que le nom de la tâche, la date d’échéance, la priorité et le statut.

Ce projet est conçu pour illustrer le fonctionnement de macros Excel dans un contexte de gestion de données type "CRUD".

---

## ⚙️ Fonctionnalité actuelle

### ✅ Ajouter une tâche (`EnregistrerTache`)

La macro `EnregistrerTache` :
- Récupère les données saisies dans la feuille **Formulaire**
- Enregistre automatiquement ces données dans la feuille **Liste des tâches**
- Vide les champs du formulaire après l’enregistrement
- Réinitialise le statut à `"À faire"`

---

## 🧪 Exemple de code (VBA)

```vba
Sub EnregistrerTache()
    Dim Form As Worksheet
    Dim Liste As Worksheet
    Dim ligne_suivant As Long

    Set Form = Sheets("Formulaire")
    Set Liste = Sheets("Liste des tâches")

    ' Trouver la prochaine ligne vide dans Liste des tâches

    ligne_suivant = Liste.Cells(Liste.Rows.Count, "F").End(xlUp).Row + 1

    ' Copier les valeurs du formulaire

    With Liste
        .Cells(ligne_suivant, "F").Value = Form.Range("H7").Value
        .Cells(ligne_suivant, "H").Value = Form.Range("H11").Value
        .Cells(ligne_suivant, "J").Value = Form.Range("H14").Value
        .Cells(ligne_suivant, "L").Value = Form.Range("H17").Value
    End With

    ' Réinitialiser les champs du formulaire

    Form.Range("H7, H11, H14").ClearContents
    Form.Range("H17").Value = "À faire"
End Sub
````
### ✅ Modifier une tâche (`MettreAJour`)

La macro `MettreAJour` :
- Récupère les données de la liste **Liste des tâches** et les mets dans **Mise à jour** afin d'effectuer les modifications.

- Il faudra sélectionner d'abord la ligne de la tâche qu'on veut modifer, ensuite cliquer sur le crayon dans **Liste des tâches**

---
## 🧪 Exemple de code (VBA)

```vba
Sub MettreAJour()
    Dim Form As Worksheet
    Dim Liste As Worksheet
    Dim ligneSelectionnee As Long

    Set Form = Sheets("Mise à jour")
    Set Liste = Sheets("Liste des tâches")

    ' Vérifie que l'utilisateur a bien sélectionné une cellule dans la colonne F

    If Not Intersect(Selection, Liste.Columns("F")) Is Nothing Then
        ligneSelectionnee = ActiveCell.Row

        ' Récupérer les valeurs de la ligne sélectionnée

        With Liste
            Form.Range("H7").Value = .Cells(ligneSelectionnee, "F").Value
            Form.Range("H10").Value = .Cells(ligneSelectionnee, "H").Value
            Form.Range("H13").Value = .Cells(ligneSelectionnee, "J").Value
            Form.Range("H16").Value = .Cells(ligneSelectionnee, "L").Value
        End With
        
        ' Stocker la ligne sélectionnée dans une cellule masquée : ce stockage permettra plus tard d'insérer les modifications dans la bonne ligne.

        Form.Range("Z1").Value = ligneSelectionnee

        ' Aller sur le formulaire de modification

        Form.Activate
    Else
        MsgBox "Veuillez sélectionner une cellule dans la colonne F (Nom de la tâche) pour mettre à jour la tâche.", vbExclamation
    End If
End Sub
````
### ✅ Enregistrer les modifications (`Update`)

La macro `Update` :
- Enregistre dans **Liste des tâches** les modifications apportées

---
## 🧪 Exemple de code (VBA)

```vba
Sub Update()

    Dim Form As Worksheet
    Dim Liste As Worksheet
    Dim ligneCible As Long

    Set Form = Sheets("Mise à jour")
    Set Liste = Sheets("Liste des tâches")

    ' Récupérer la ligne d'origine à modifier

   ligneCible = Form.Range("Z1").Value

    ' Vérification
     If ligneCible <= 0 Then
       MsgBox "Aucune ligne de tâche sélectionnée pour la mise à jour.", vbExclamation
        Exit Sub
    End If
    
    ' Mise à jour des données dans la ligne cible

    With Liste
        .Cells(ligneCible, "F").Value = Form.Range("H7").Value
        .Cells(ligneCible, "H").Value = Form.Range("H10").Value
        .Cells(ligneCible, "J").Value = Form.Range("H13").Value
        .Cells(ligneCible, "L").Value = Form.Range("H16").Value
    End With

    
    ' Nettoyer les champs 

    Form.Range("H7, H10, H13, H16, Z1").ClearContents
    
    MsgBox "Tâche mise à jour avec succès !", vbInformation


End Sub
````

### ✅ Supprimer une tâche(`Supprimer_tâche`)

La macro `Supprimer_tâche` :
- Supprime une tâche dans **Liste des tâches**
- Il faudra sélectionner d'abord la ligne de la tâche qu'on veut supprimer, ensuite cliquer sur l'icône de la poubelle pour supprimer

---
## 🧪 Exemple de code (VBA)

```vba
Sub Supprimer_tâche()


    Dim Liste As Worksheet
    Dim ligneSelectionnee As Long

    Set Liste = Sheets("Liste des tâches")
    
    ligneSelectionnee = ActiveCell.Row
    
    Liste.Rows(ligneSelectionnee).Delete

End Sub
````
* Enregistrer un tache
![image](https://github.com/user-attachments/assets/aad80aa6-27b0-478b-9abd-98c2b99a92c7)

* Liste des taches
![image](https://github.com/user-attachments/assets/81df36a2-f608-4bca-948c-c96da5724fe6)

* Mise à jour
![image](https://github.com/user-attachments/assets/53af4b8a-183d-406c-8a21-4ff503657dd4)

* Suppression de la tache
![image](https://github.com/user-attachments/assets/749e21d1-e95a-4edf-8321-87c650cc9500)




