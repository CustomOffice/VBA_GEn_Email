# VBA_Gen_Email
Macro pour générer un email avec outlook en 3 partie, un texte (intro), un tableau et un texte (signature)			

##Lien vers le site
http://customoffice.github.io/VBA_Gen_Email/

## Instruction
- Soit créer un module dans votre projet vba et y copier/coller le code ci-dessous
- Soit télécharger le module (fichier *.bas) et l'inserer dans votre projet vba

##Code
```bash
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!TITRE : Génération d'email avec Outlook                                                         !!!
'!!!DATE : 17.04.2015                                                                               !!!
'!!!                                                                                                !!!
'!!!DESCRIPTION : Macro pour générer un email avec outlook en 3 partie, un texte, un tableau et un 	!!!
'!!!texte													                                    	!!!
'!!!                                                                                                !!!
'!!!REGLES :                                                                                        !!!
'!!!- intro, tableau et signature sont des tableaux de valeurs, avec une valeur par retour à ligne  !!!
'!!!- tableau représente le tableau présent dans le mail, ils feront les mêmes dimensions           !!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Sub gen_email_tabl(destinataire As String, objet As String, ByRef intro, ByRef tableau, ByRef signature, titre_tableau As Boolean)
    'Déclaration des variables
        'les objets
    Dim objOL As Object, olmail As Object
        'les constantes
    Const olmailItem = 0 ' Outlook VBA constant olMailItem
        'les string
    Dim content_mail As String, style_bordure_tableau As String, police As String, couleur_police As String
    Dim couleur_bordure As String, style_mail As String
        'les integer
    Dim t As Integer, l As Integer, i As Integer, taille_police As Integer, epaisseur_bordure As Integer, marge_interieur_cellule As Integer
    
    'Paramètre de personnalisation du mail
    police = "Arial" ' n'utilisez que des polices "standard"
    taille_police = 14 'en px
    couleur_police = "RGB(31,73,125)" 'couleur en RGB, entre le blanc 255,255,255 et le noir 0,0,0
    epaisseur_bordure = 1 'en px
    couleur_bordure = "RGB(31,73,125)" 'couleur en RGB, entre le blanc 255,255,255 et le noir 0,0,0
    marge_interieur_cellule = 5 'en px
    
    'création des styles
    style_bordure_tableau = "style='font-family:" & police & ";font-size:" & taille_police & "px;border:" & epaisseur_bordure & "px solid " & couleur_bordure & ";padding:" & marge_interieur_cellule & "px; color:" & couleur_police & "'"
    style_mail = "style='font-family:" & police & ";font-size:" & taille_police & "px;color:" & couleur_police & "'"
    
    'init du contenu html du mail
    content_mail = "<body " & style_mail & ">"
    content_mail = content_mail & "<p>"
    
    'ajout du texte d'intro du mail
    For i = 0 To UBound(intro)
        content_mail = content_mail & intro(i) & "<br/>"
    Next i
    
    'init du tableau
    content_mail = content_mail & "<table style='border-collapse:collapse'>"
    
    'ajout des titres du tableau si existant dans la variable
    If titre_tableau = True Then
        content_mail = content_mail & "<tr>"
        For i = 0 To UBound(tableau, 2)
            content_mail = content_mail & "<th " & style_bordure_tableau & ">" & tableau(0, i) & "</th>"
        Next i
        content_mail = content_mail & "</tr>"
        t = 1 'permet de savoir à quelle ligne du tableau commencer, 1 s'il y a des titres sinon 0
    Else
        t = 0
    End If
    
    'ajout du reste du tableau
    For l = t To UBound(tableau, 1)
        content_mail = content_mail & "<tr>"
        For i = 0 To UBound(tableau, 2)
            content_mail = content_mail & "<td " & style_bordure_tableau & ">" & tableau(l, i) & "</td>"
        Next i
        content_mail = content_mail & "</tr>"
    Next l
    
    'fermeture du tableau
    content_mail = content_mail & "</table>"
    
    'ajout de la signature
    For i = 0 To UBound(signature)
        content_mail = content_mail & signature(i) & "<br/>"
    Next i
    
    'fermeture du contenu du mail
    content_mail = content_mail & "</p></body>"
        
    'création du mail
    Set objOL = CreateObject("Outlook.Application")
    Set olmail = objOL.CreateItem(olmailItem)
    With olmail
        .To = destinataire
        .Subject = objet
        .HTMLBody = content_mail
        .Display
    End With

End Sub
```
