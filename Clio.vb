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
        
        'Idem, Id., Ibidem et Ibid. en italiques
         With Selection.Find
            .ClearFormatting
            .Text = "Ibidem"
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
            .Text = "Ibid."
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
            .Text = "Idem"
            .Wrap = wdFindContinue
            .MatchWholeWord = True
            .MatchWildcards = False
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
            .Text = "Id."
            .Wrap = wdFindContinue
            .MatchWholeWord = True
            .MatchWildcards = False
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

Sub Clio_2_regnes()
' Permet d'ajouter une espace insécable entre le nom d'un souverain et son numéro dynastique
' Gère actuellement les souverains suivants (ordre alphabétique) : Catherine, Charles, Édouard,
' Edward, François, Henri, Jean, Jules, Léon, Louis, Napoléon, Richard.
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
        
    'Surligner particule / article et patronyme
        ' de/du + Nom
            With Selection.Find
            .ClearFormatting
            .Text = "<d([eu]) ([A-ZÉ])"
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

Sub Clio_3_numbers()
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

Sub Clio_4_stylage()

   'Stylage caractères italique
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Italic = True
        .SmallCaps = False
        .AllCaps = False
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
        .AllCaps = False
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
    
' Surligner les guillements ouvrants
        With Selection.Find
            .ClearFormatting
            .Forward = True
            .Wrap = wdFindContinue
            .Text = "(«)^s"
            .MatchWildcards = True
            With .Replacement
                .ClearFormatting
                .Text = ""
                .Highlight = True
            End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End With
        
' Surligner les guillements fermants
        With Selection.Find
            .ClearFormatting
            .Forward = True
            .Wrap = wdFindContinue
            .Text = "^s(»)"
            .MatchWildcards = True
            With .Replacement
                .ClearFormatting
                .Text = ""
                .Highlight = True
            End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End With
        

End Sub

Sub Clio_5_nettoyage()
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
            End With
        End With
End Sub

