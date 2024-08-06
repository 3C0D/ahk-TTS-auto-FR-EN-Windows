#Requires AutoHotkey v2.0

; Variables globales
global isReading := false
global voice := ComObject("SAPI.SpVoice")

; Définit le raccourci Windows+Y pour l'autodétection
#y:: ReadText("AUTO")

ReadText(language) {
    global isReading, voice

    if (voice.Status.RunningState == 2) {
        voice.Speak("", 3)  ; Arrête la lecture en cours
        isReading := false
        return
    }

    if (isReading) {
        isReading := false
    }

    OldClipboard := A_Clipboard
    A_Clipboard := ""
    
    Send "^c"  ; Copie le texte sélectionné
    if !ClipWait(2) {
        if (OldClipboard != "") {
            SelectedText := OldClipboard
            A_Clipboard := OldClipboard
        } else {
            ; Si ni sélection ni presse-papiers retourner
            MsgBox "Aucun texte sélectionné ou contenu dans le presse-papiers"
            A_Clipboard := OldClipboard
            return
        }
    } else {
        ; Si du texte est sélectionné, l'utiliser pour la synthèse vocale
        SelectedText := A_Clipboard
    }
    
    SelectedText := IgnoreCharacters(SelectedText)
    
    try {
        SetVoiceLanguage(language, SelectedText)
        voice.Rate := 2
        
        isReading := true
        voice.Speak(SelectedText, 1)  ; Lecture asynchrone
    } catch as err {
        MsgBox "Erreur lors de l'utilisation de la synthèse vocale: " . err.Message
        isReading := false
    }
    
    A_Clipboard := OldClipboard
}

SetVoiceLanguage(language, text := "") {
    if (language == "AUTO") {
        language := DetectLanguage(text)
    }

    ; Utiliser les noms exacts des voix disponibles
    if (language == "EN") {
        voiceName := "Microsoft Zira Desktop"
    } else if (language == "FR") {
        voiceName := "Microsoft Hortense Desktop"  ; Utilisez Hortense pour le français
    } else {
        MsgBox "Langue non supportée : " . language
        return
    }

    for v in ComObject("SAPI.SpVoice").GetVoices() {
        if (v.GetAttribute("Name") == voiceName) {
            voice.Voice := v
            return
        }
    }
    
    MsgBox "Voix pour la langue " . language . " non trouvée. Utilisation de la voix par défaut."
}

DetectLanguage(text) {
    ; Détection de la langue basée sur des mots courants
    frenchWords := ["le", "la", "les", "un", "une", "des", "et", "ou", "mais", "donc", "or", "ni", "car", "que", "qui", "quoi", "dont", "où", "à", "au", "avec", "pour", "sur", "dans", "par", "ce", "cette", "ces"]
    englishWords := ["the", "and", "or", "but", "so", "yet", "for", "nor", "that", "which", "who", "whom", "whose", "when", "where", "why", "how", "a", "an", "in", "on", "at", "with", "by", "this", "these", "those", "is"]

    frenchScore := 0
    englishScore := 0

    words := StrSplit(text, " ")
    for word in words {
        if (HasVal(frenchWords, word))
            frenchScore++
        if (HasVal(englishWords, word))
            englishScore++
    }

    if (englishScore > frenchScore) {
        return "EN"
    } else if (frenchScore > englishScore) {
        return "FR"
    } else {
        return "FR"  ; Par défaut, considère le français
    }
}

HasVal(haystack, needle) {
    for index, value in haystack
        if (value = needle)
            return true
    return false
}

IgnoreCharacters(text) {
    charactersToIgnore := ["*", "#", "@"]
    for char in charactersToIgnore {
        text := StrReplace(text, char, "")
    }
    return text
}
