Sub Clio_typographie()
' Gère les éléments suivants :
' 1. la normalisation de la typographie sur des erreurs communes ;
' 2. l'ajout de surlignement sur les éléments qui demandent une vérification humaine ;
' 3. l'ajout d'espaces insécables dans des expressions régulières - abréviations, bibliographie.

    ' Normalisation de la typographie
        ' Remplacement faux points de suspension
    With Selection.Find
        .ClearFormatting
        .Text = "..."
        .Wrap = wdFindContinue
        With .Replacement
            .ClearFormatting
            .Text = "…"
        End With
        .Execute Replace:=wdReplaceAll
    End With
        
        ' Suppression de la suite virgule - points de suspension
    With Selection.Find
        .Text = ",…"
        .Wrap = wdFindContinue
        With .Replacement
            .Text = "…"
            End With
        .Execute Replace:=wdReplaceAll
        End With
        
        'Suppression de la suite virgule - espace - points de suspension
    With Selection.Find
        .Text = ",^w…"
        .Wrap = wdFindContinue
        With .Replacement
            .Text = "…"
            End With
        .Execute Replace:=wdReplaceAll
        End With
    
        ' Retirer espaces avant ponctuation simple
    With Selection.Find
        .Text = " {1;}([,.…”])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Ajout espace insécable avant signe de ponctuation haute
    With Selection.Find
        .ClearFormatting
        .Text = " {1;}([\!\?»:;%])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "^s\1"
            .Highlight = True
            End With
        .Execute Replace:=wdReplaceAll
    End With
    
    
    ' Alerter sur de potentielles erreurs
        ' Surligner espaces entre apostrophe et mot suivant
    With Selection.Find
        .ClearFormatting
        .Text = "([a-zA-Z])’ {1;}([a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = ""
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
        'Surligner les espaces insécables avant un caractère alphanumérique
    With Selection.Find
        .ClearFormatting
        .Text = "^s([0-9a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = ""
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
            ' Sauf si précédé de guillemets français ouvrant
            With Selection.Find
                .Text = "«^s([0-9a-zA-Z])"
                .Wrap = wdFindContinue
                .MatchWildcards = True
                With .Replacement
                    .Text = ""
                    .Highlight = False
                End With
            .Execute Replace:=wdReplaceAll
            End With
            ' Sauf si précédé des abréviations n. p. t. v. + insécable
            With Selection.Find
                .ClearFormatting
                .Text = "(<[nptv]).^s([1-9a-zA-Z])"
                .Wrap = wdFindContinue
                .MatchWildcards = True
                With .Replacement
                    .Text = ""
                    .Highlight = False
                End With
            .Execute Replace:=wdReplaceAll
            End With
            'Sauf si suit une initiale majuscule
             With Selection.Find
                .ClearFormatting
                .Text = "([A-ZÉ])^s.([A-ZÉ])"
                .Wrap = wdFindContinue
                .MatchWildcards = True
                With .Replacement
                    .Text = ""
                    .Highlight = False
                End With
                .Execute Replace:=wdReplaceAll
            End With
                    
        ' Surligner tirets cadratins et demi-cadratins
        With Selection.Find
            .ClearFormatting
            .Text = "([—–])"
            .Wrap = wdFindContinue
            With .Replacement
                .Text = "\1"
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
        
        ' Surligner "Age" - pas de correction automatique à cause des textes en anglais
        With Selection.Find
            .ClearFormatting
            .Text = "Age"
            .Wrap = wdFindContinue
            .MatchCase = True
            .MatchWildcards = False
            With .Replacement
                .Text = ""
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
    
    ' Gestions des insécables après des abréviations
        ' Espaces insécables après initiale majuscule
    With Selection.Find
        .ClearFormatting
        .Text = "([A-ZÉ]). "
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1.^s"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
        ' Ajout espaces et insécable entre n., p., t., v. et numéro ou lettre
    With Selection.Find
        .ClearFormatting
        .Text = "<([nptv]).([1-9a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1. \2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(<[nptv]). ([1-9a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1.^s\2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
        ' Ajout espace et insécable entre abréviation éd. et lettre – FR
    With Selection.Find
        .ClearFormatting
        .Text = "<(éd.)([A-ZÉ])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1 \2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "<(éd.) ([A-ZÉ])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With

        ' Ajout espace et insécable entre abréviation ed. et lettre – EN
    With Selection.Find
        .ClearFormatting
        .Text = "<(ed.)([A-ZÉ])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1 \2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "<(ed.) ([A-ZÉ])"
        .Wrap = wdFindContinue
        .MatchWildcards = False
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
        ' Ajout espace et insécable entre abréviation dir et lettre
    With Selection.Find
        .ClearFormatting
        .Text = "<(dir.)([A-ZÉ])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1 \2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "<(dir.) ([A-ZÉ])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
        'Ajout espace et insécable entre abréviation vol et lettre ou  numéro
    With Selection.Find
        .ClearFormatting
        .Text = "<(vol.)([0-9a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1 \2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "<(vol.) ([0-9a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Ajout d'insécables dans certaines expressions
        ' Ajout espace insécable entre e exposant et siècle
    With Selection.Find
        .ClearFormatting
        .Text = "e siècle"
        .Wrap = wdFindContinue
        .MatchWildcards = False
        With .Replacement
            .Text = "e^ssiècle"
            .Font.Superscript = True
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "^ssiècle"
        .Font.Superscript = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = ""
            .Font.Superscript = False
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
        .Replacement.ClearFormatting
    End With
    
        ' Ajout espace insécable dans la suite <nombre e éd.>
    With Selection.Find
        .ClearFormatting
        .Text = "([0-9]{1;})e éd."
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .ClearFormatting
            .Text = "\1e^séd."
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "e^séd."
        .Wrap = wdFindContinue
        .MatchWildcards = False
        With .Replacement
            .Text = "e^séd."
            .Font.Superscript = True
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "^séd."
        .Font.Superscript = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = ""
            .Font.Superscript = False
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
        .Replacement.ClearFormatting
    End With
    
        ' Normalisation op. cit.
    With Selection.Find
        .ClearFormatting
        .Text = "op. cit,"
        .Wrap = wdFindContinue
        .MatchWildcards = False
        With .Replacement
            .Text = "op. cit.,"
            .Font.Italic = False
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
        .Replacement.ClearFormatting
    End With   
    With Selection.Find
        .ClearFormatting
        .Text = "op.cit."
        .Wrap = wdFindContinue
        .MatchWildcards = False
        With .Replacement
            .Text = "op. cit."
            .Highlight = False
            End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End With
    With Selection.Find
            .ClearFormatting
            .Text = "op. cit."
            .Wrap = wdFindContinue
            .MatchWildcards = False
        With .Replacement
            .Text = "op.^scit."
            .Font.Italic = True
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
        .Replacement.ClearFormatting
    End With
    
        ' Ajouter espace insécable après Cf ou cf.
    With Selection.Find
        .ClearFormatting
        .Text = "([cC])f. "
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1f.^s"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
    'Nettoyer en partant
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .MatchWildcards = False
            With .Replacement
                .Text = ""
                .ClearFormatting
                .MatchWildcards = False
            End With
        End With

End Sub

Sub Clio_regnes()
' Permet d'ajouter une espace insécable entre le nom d'un souverain et son rang dynastique
' Gère actuellement les souverains suivants (ordre alphabétique) : Catherine, Charles, Édouard,
' Edward, François, Henri, Jean, Jules, Léon, Louis, Napoléon, Richard. 
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        'Catherine
            With Selection.Find
            .ClearFormatting
            .Text = "(Catherine) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Charles
            With Selection.Find
            .ClearFormatting
            .Text = "(Charles) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Édouard
            With Selection.Find
            .ClearFormatting
            .Text = "(Édouard) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Edward
            With Selection.Find
            .ClearFormatting
            .Text = "(Edward) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
        'François
            With Selection.Find
            .ClearFormatting
            .Text = "(François) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Henri
            With Selection.Find
            .ClearFormatting
            .Text = "(Henri) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
         'Jean
            With Selection.Find
            .ClearFormatting
            .Text = "(Jean) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Jules
            With Selection.Find
            .ClearFormatting
            .Text = "(Jules) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Léon
            With Selection.Find
            .ClearFormatting
            .Text = "(Léon) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        ' Louis
            With Selection.Find
            .ClearFormatting
            .Text = "(Louis) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Napoléon
            With Selection.Find
            .ClearFormatting
            .Text = "(Napoléon) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Richard
            With Selection.Find
            .ClearFormatting
            .Text = "(Richard) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With

    'Nettoyer en partant
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .MatchWildcards = False
            With .Replacement
                .Text = ""
                .ClearFormatting
                .MatchWildcards = False
            End With
        End With

End Sub

Sub Clio_numbers()
' Permet de surligner tous les nombres du document pour faciliter l'ajout manuel d'insécables :
' nombres longs, dates, etc.
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = True
    
        ' Surligner tous les chiffres arabes
        With Selection.Find
            .Text = "^#"
            .MatchWildcards = False
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll

        ' Surligner les capitales utilisées pour les chiffres romains
        With Selection.Find
            .Text = "([CDILMVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1"
                .Highlight = True
            End With
        .Execute Replace:=wdReplaceAll
        End With
        
        'Nettoyer en partant
            With Selection.Find
                .ClearFormatting
                .Text = ""
                .MatchWildcards = False
                With .Replacement
                    .Text = ""
                    .ClearFormatting
                End With
            End With
End Sub

Sub Clio_nettoyage()
' Supprime le surlignage dans le document
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .Text = ""
        .MatchWildcards = False
        .Replacement.Text = ""
        .Wrap = wdFindContinue
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    'Nettoyer en partant
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .MatchWildcards = False
            With .Replacement
                .Text = ""
                .ClearFormatting
                .MatchWildcards = False
            End With
        End With
End Sub