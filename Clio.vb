Sub Clio_1_typographie()
' Gère les éléments suivants :
' 1. la normalisation de la typographie sur des erreurs communes ;
' 2. l'ajout de surlignement sur les éléments qui demandent une vérification humaine ;
' 3. l'ajout d'espaces insécables dans des expressions régulières - abréviations, bibliographie.

    ' Normalisation de la typographie
        ' Supprimer deux marques de paragraphe successives
    With Selection.Find
        .ClearFormatting
        .Text = "^p^p"
        .Wrap = wdFindContinue
        With .Replacement
            .ClearFormatting
            .Text = "^p"
        End With
        .Execute Replace:=wdReplaceAll
    End With
        
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
        ' Surligner deux apostrophes au lieu de guillemets anglais
    With Selection.Find
        .ClearFormatting
        .Text = "''"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = ""
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
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
        .Text = "^s([0-9a-zA-Z\(\[])"
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
                .Text = "«^s([0-9a-zA-Z\(\[])"
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
                .Text = "(<[nptv]).^s([0-9a-zA-Z])"
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
                            
        'Surligner le signe degré
        With Selection.Find
            .ClearFormatting
            .Text = "°"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            With .Replacement
                .Text = ""
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
        
        ' Surligner tirets cadratins et demi-cadratins
        With Selection.Find
            .ClearFormatting
            .Text = "([—–])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
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
        
        ' Surligner les appels de notes précédés d'une espace
        With Selection.Find
            .ClearFormatting
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Text = " ^f"
            With .Replacement
                .Text = ""
                .Highlight = True
            End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End With

        ' Surligner les appels de notes
        With Selection.Find
            .ClearFormatting
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Text = "^f"
            With .Replacement
                .Text = ""
                .Highlight = True
            End With
        Selection.Find.Execute Replace:=wdReplaceAll
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
    
        ' Espaces insécables après initiale + h + majuscule
    With Selection.Find
        .ClearFormatting
        .Text = "([A-Z]h). "
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
    
        ' Ajout espace et insécable entre abréviation art. et lettre ou numéro
    With Selection.Find
        .ClearFormatting
        .Text = "<(art.)([0-9a-zA-Z])"
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
        .Text = "<(art.) ([0-9a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With

        ' Ajout espace et insécable entre abréviation chap. et lettre ou numéro
    With Selection.Find
        .ClearFormatting
        .Text = "<(chap.)([0-9a-zA-Z])"
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
        .Text = "<(chap.) ([0-9a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
        
        ' Ajout espace et insécable entre abréviation coll. et lettre ou numéro
    With Selection.Find
        .ClearFormatting
        .Text = "<(coll.) "
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s"
            .Highlight = True
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
    
        ' Ajout espace et insécable entre abréviation liv. et lettre ou numéro
    With Selection.Find
        .ClearFormatting
        .Text = "<(liv.)([0-9a-zA-Z])"
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
        .Text = "<(liv.) ([0-9a-zA-Z])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With
        
        ' Ajout espace et insécable entre abréviation part. et lettre ou numéro
    With Selection.Find
        .ClearFormatting
        .Text = "<(part.)([0-9a-zA-Z])"
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
        .Text = "<(part.) ([0-9a-zA-Z])"
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
            .Highlight = True
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
            .Highlight = True
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
 
        ' Remplacement n° en no sup + insécable
        With Selection.Find
            .ClearFormatting
            .Text = "n° "
            .Wrap = wdFindContinue
            .MatchWildcards = False
            With .Replacement
                .Text = "n°"
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
            With Selection.Find
            .ClearFormatting
            .Text = "n°^s"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            With .Replacement
                .Text = "n°"
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "n°"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            With .Replacement
                .Text = "no "
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
            With Selection.Find
            .ClearFormatting
            .Text = "<(no)([0-9])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1 \2"
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "no "
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            With .Replacement
                .Text = "no^s"
                .Highlight = True
                .Font.Superscript = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "n"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            .Font.Superscript = True
            With .Replacement
                .Text = "n"
                .Highlight = True
                .Font.Superscript = False
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "^s"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            .Font.Superscript = True
            With .Replacement
                .Text = "^s"
                .Font.Superscript = False
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        
        ' Remplacement f° en fo sup
        With Selection.Find
            .ClearFormatting
            .Text = "f°"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            With .Replacement
                .Text = "fo"
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "fo"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            With .Replacement
                .Text = ""
                .Highlight = True
                .Font.Superscript = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "f"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            .Font.Superscript = True
            With .Replacement
                .Text = "f"
                .Highlight = True
                .Font.Superscript = False
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        
        ' Remplacement r° en ro sup
        With Selection.Find
            .ClearFormatting
            .Text = "r°"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            With .Replacement
                .Text = "ro"
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "ro"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            With .Replacement
                .Text = ""
                .Highlight = True
                .Font.Superscript = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "r"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            .Font.Superscript = True
            With .Replacement
                .Text = "r"
                .Highlight = True
                .Font.Superscript = False
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        
        ' Remplacement v° en vo sup
        With Selection.Find
            .ClearFormatting
            .Text = "v°"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            With .Replacement
                .Text = "vo"
                .Highlight = True
            End With
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "vo"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            With .Replacement
                .Text = ""
                .Highlight = True
                .Font.Superscript = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "v"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Highlight = True
            .Font.Superscript = True
            With .Replacement
                .Text = "v"
                .Highlight = True
                .Font.Superscript = False
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        
        
        'Idem, Id., Ibidem et Ibid. en italiques
         With Selection.Find
            .ClearFormatting
            .Text = "<([Ii]bidem)"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .MatchWholeWord = True
            .Font.Italic = False
            With .Replacement
                .Text = ""
                .Font.Italic = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "<([Ii]bid.)"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .MatchWholeWord = True
            .Font.Italic = False
            With .Replacement
                .Text = ""
                .Font.Italic = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "<([Ii]dem)"
            .Wrap = wdFindContinue
            .MatchWholeWord = True
            .MatchWildcards = True
            .Font.Italic = False
            With .Replacement
                .Text = ""
                .Font.Italic = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
                With Selection.Find
            .ClearFormatting
            .Text = "<([Ii]d.)"
            .Wrap = wdFindContinue
            .MatchWholeWord = True
            .MatchWildcards = True
            .Font.Italic = False
            With .Replacement
                .Text = ""
                .Font.Italic = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "supra"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .MatchWholeWord = True
            .Font.Italic = False
            With .Replacement
                .Text = ""
                .Font.Italic = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
        With Selection.Find
            .ClearFormatting
            .Text = "infra"
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .MatchWholeWord = True
            .Font.Italic = False
            With .Replacement
                .Text = ""
                .Font.Italic = True
            End With
            .Execute Replace:=wdReplaceAll
            .Replacement.ClearFormatting
        End With
           
    'Nettoyer en partant
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .MatchWildcards = False
            .MatchWholeWord = False
            With .Replacement
                .Text = ""
                .ClearFormatting
            End With
        End With
End Sub

Sub Clio_2_nomspropres()
' Permet d'ajouter une espace insécable entre le nom d'un souverain et son numéro dynastique
' Gère actuellement les souverains suivants (ordre alphabétique) : Catherine, Charles, Édouard,
' Edward, François, Henri, Innocent, Jean, Jules, Léon, Louis, Napoléon, Richard.
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        'Albert
            With Selection.Find
            .ClearFormatting
            .Text = "(Albert) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        'Alexandre
            With Selection.Find
            .ClearFormatting
            .Text = "(Alexandre) ([IVX])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
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
            End With
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
        'Henri
            With Selection.Find
            .ClearFormatting
            .Text = "(Innocent) ([IVX])"
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
        'Philippe
            With Selection.Find
            .ClearFormatting
            .Text = "(Philippe) ([IVX])"
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
        
    'Surligner particule / article et patronyme
        ' de/du + Nom
            With Selection.Find
            .ClearFormatting
            .Text = "<([dD])([eu]) ([A-ZÉ])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = ""
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With

        ' de La + Nom
            With Selection.Find
            .ClearFormatting
            .Text = "de ([lL])a ([A-ZÉ])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = ""
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
            
        ' de L' + Nom
            With Selection.Find
            .ClearFormatting
            .Text = "de ([lL])(?)([A-ZÉ])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            With .Replacement
                .Text = ""
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
            
            
        ' La/Le + Nom
            With Selection.Find
            .ClearFormatting
            .Text = "<L([ae]) ([A-ZÉ])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .MatchCase = True
            With .Replacement
                .Text = ""
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
    
    'Ajouter insécable entre particule étrangère et patronyme
        
        ' van + Nom
            With Selection.Find
            .ClearFormatting
            .Text = "<([vV]an) ([A-ZÉ])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .MatchCase = False
            With .Replacement
                .Text = "\1^s\2"
                .Highlight = True
                End With
            .Execute Replace:=wdReplaceAll
            End With
        
        ' von + Nom
            With Selection.Find
            .ClearFormatting
            .Text = "<([vV]on) ([A-ZÉ])"
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .MatchCase = False
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
            .MatchCase = False
            With .Replacement
                .Text = ""
                .ClearFormatting
            End With
        End With
End Sub
Sub Clio_3_dates()
' espacement dates moi année
    ' janvier
    With Selection.Find
        .ClearFormatting
        .Text = "(janvier)^s([0-9])"
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
        .Text = "([0-9]) (janvier)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (janvier)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (janvier) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' février
    With Selection.Find
        .ClearFormatting
        .Text = "(février)^s([0-9])"
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
        .Text = "([0-9]) (février)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (février)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (février) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' mars
    With Selection.Find
        .ClearFormatting
        .Text = "(mars)^s([0-9])"
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
        .Text = "([0-9]) (mars)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (mars)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (mars) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' avril
    With Selection.Find
        .ClearFormatting
        .Text = "(avril)^s([0-9])"
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
        .Text = "([0-9]) (avril)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (avril)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (avril) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' mai
    With Selection.Find
        .ClearFormatting
        .Text = "(mai)^s([0-9])"
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
        .Text = "([0-9]) (mai)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (mai)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (mai) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' juin
    With Selection.Find
        .ClearFormatting
        .Text = "(juin)^s([0-9])"
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
        .Text = "([0-9]) (juin)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (juin)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (juin) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' juillet
    With Selection.Find
        .ClearFormatting
        .Text = "(juillet)^s([0-9])"
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
        .Text = "([0-9]) (juillet)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (juillet)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (juillet) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' août
    With Selection.Find
        .ClearFormatting
        .Text = "(août)^s([0-9])"
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
        .Text = "([0-9]) (août)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (août)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (août) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' septembre
    With Selection.Find
        .ClearFormatting
        .Text = "(septembre)^s([0-9])"
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
        .Text = "([0-9]) (septembre)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (septembre)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (septembre) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' octobre
    With Selection.Find
        .ClearFormatting
        .Text = "(octobre)^s([0-9])"
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
        .Text = "([0-9]) (octobre)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (octobre)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (octobre) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' novembre
    With Selection.Find
        .ClearFormatting
        .Text = "(novembre)^s([0-9])"
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
        .Text = "([0-9]) (novembre)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (novembre)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (novembre) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    ' décembre
    With Selection.Find
        .ClearFormatting
        .Text = "(décembre)^s([0-9])"
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
        .Text = "([0-9]) (décembre)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "(er) (décembre)"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        With .Replacement
            .Text = "\1^s\2"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
    With Selection.Find
        .ClearFormatting
        .Text = "([a-z]) (décembre) ([0-9])"
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Highlight = False
        With .Replacement
            .Text = "\1 \2^s\3"
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub Clio_4_numbers()
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
            ' Supprimer le surlignement sur ces capitales si suivies de minuscules (sauf e)
            With Selection.Find
                .Text = "([CDILMVX])([abcdfghijklmnopqrstuvwxyzàâéèêîôû]{1;})"
                .Wrap = wdFindContinue
                .MatchWildcards = True
                With .Replacement
                    .Text = ""
                    .Highlight = False
                End With
            .Execute Replace:=wdReplaceAll
            End With
            'Supprimer le surlignement sur ces capitales suivies de e + autre lettre
            With Selection.Find
                .Text = "([CDILMVX])e([a-z]{1;})"
                .Wrap = wdFindContinue
                .MatchWildcards = True
                With .Replacement
                    .Text = ""
                    .Highlight = False
                End With
            .Execute Replace:=wdReplaceAll
            End With
            
            'Supprimer le surlignement sur ces capitales suivies d'une apostrophe
             With Selection.Find
                .Text = "([CDILMVX])(^0146)"
                .Wrap = wdFindContinue
                .MatchWildcards = True
                With .Replacement
                    .Text = ""
                    .Highlight = False
                End With
            .Execute Replace:=wdReplaceAll
            End With
        
        'Surligner les petites capitales
        With Selection.Find
            .Text = ""
            .Wrap = wdFindContinue
            .MatchWildcards = False
            .Font.SmallCaps = True
            With .Replacement
                .Text = ""
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


Sub Clio_5_createstyles()
'Permet de créer des styles de caractère dans le document pour remplacer la mise en forme directe

    ' Bold
    Set myStyle = ActiveDocument.Styles.Add(Name:="_bold", _
     Type:=wdStyleTypeCharacter)
    With myStyle.Font
     .Bold = True
     .Color = RGB(255, 147, 0)
    End With
    
    ' Italique
    Set myStyle = ActiveDocument.Styles.Add(Name:="_italic", _
     Type:=wdStyleTypeCharacter)
    With myStyle.Font
     .Italic = True
     .Color = RGB(255, 147, 0)
    End With
    
    ' Smallcaps-i
    Set myStyle = ActiveDocument.Styles.Add(Name:="_smallcaps-i", _
     Type:=wdStyleTypeCharacter)
    With myStyle.Font
     .Italic = True
     .SmallCaps = True
     .Color = RGB(255, 147, 0)
    End With

    ' Smallcaps
    Set myStyle = ActiveDocument.Styles.Add(Name:="_smallcaps", _
     Type:=wdStyleTypeCharacter)
    With myStyle.Font
     .SmallCaps = True
     .Color = RGB(255, 147, 0)
    End With

    ' Sup
    Set myStyle = ActiveDocument.Styles.Add(Name:="_sup", _
     Type:=wdStyleTypeCharacter)
    With myStyle.Font
     .Superscript = True
     .Color = RGB(255, 147, 0)
    End With

    ' Sup-i
    Set myStyle = ActiveDocument.Styles.Add(Name:="_sup-i", _
     Type:=wdStyleTypeCharacter)
    With myStyle.Font
     .Italic = True
     .Superscript = True
     .Color = RGB(255, 147, 0)
    End With
    
    ' Underline
    Set myStyle = ActiveDocument.Styles.Add(Name:="_underline", _
     Type:=wdStyleTypeCharacter)
    With myStyle.Font
     .Underline = True
     .Color = RGB(255, 147, 0)
    End With
    
End Sub

Sub Clio_6_stylage()
' Permet de remplacer la mise en forme directe du document par un stylage de caractères.
' Nécessite ensuite de passer en revue le document pour repérer les éléments qui ne devraient pas avoir de mise en forme particulière.

   'Stylage caractères bold
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .Italic = False
        .SmallCaps = False
        .Superscript = False
        .Subscript = False
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("_bold")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'Stylage caractères italique
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Italic = True
        .SmallCaps = False
        .Superscript = False
        .Subscript = False
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("_italic")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'Stylage caractères smallcaps-i
        Selection.Find.ClearFormatting
        With Selection.Find.Font
            .Italic = True
            .SmallCaps = True
            .Superscript = False
            .Subscript = False
        End With
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("_smallcaps-i")
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    
    'Stylage caractères smallcaps
        Selection.Find.ClearFormatting
        With Selection.Find.Font
            .Italic = False
            .SmallCaps = True
            .AllCaps = False
            .Superscript = False
            .Subscript = False
        End With
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("_smallcaps")
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    
    'Stylage caractères sup
        Selection.Find.ClearFormatting
        With Selection.Find.Font
            .Italic = False
            .SmallCaps = False
            .AllCaps = False
            .Superscript = True
            .Subscript = False
        End With
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("_sup")
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    
    'Stylage caractères sup-i
        Selection.Find.ClearFormatting
        With Selection.Find.Font
            .Italic = True
            .SmallCaps = False
            .AllCaps = False
            .Superscript = True
            .Subscript = False
        End With
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("_sup-i")
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        'Stylage caractères underline
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = False
        .Italic = False
        .Underline = True
        .SmallCaps = False
        .Superscript = False
        .Subscript = False
    End With
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("_underline")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'Stylage paragraphes Corps de texte > Normal
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("Corps de texte")
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    ' Surligner les guillements ouvrants
            'With Selection.Find
                '.ClearFormatting
                '.Forward = True
                '.Wrap = wdFindContinue
                '.Text = "(«)^s"
                '.MatchWildcards = True
                'With .Replacement
                    '.ClearFormatting
                    '.Text = ""
                    '.Highlight = True
                'End With
            'Selection.Find.Execute Replace:=wdReplaceAll
            'End With
            
    ' Surligner les guillements fermants
            'With Selection.Find
                '.ClearFormatting
                '.Forward = True
                '.Wrap = wdFindContinue
                '.Text = "^s(»)"
                '.MatchWildcards = True
                'With .Replacement
                    '.ClearFormatting
                    '.Text = ""
                    '.Highlight = True
                'End With
            'Selection.Find.Execute Replace:=wdReplaceAll
            'End With

    ' Nettoyer en partant
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

Sub Clio_7_nettoyage()
' Pour nettoyer entre chaque macro
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
            .MatchWholeWord = False
            .MatchCase = False
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
            End With
        End With
End Sub



