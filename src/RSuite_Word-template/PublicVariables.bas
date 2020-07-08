Attribute VB_Name = "PublicVariables"
Public OPTHYPH, OPTHYPH2, NBHYPH, SOQ, SCQ, DOQ, DCQ, DP, SP, aSPACE, RTN, RTN2, EMDASH, ENDASH, ELLIPSIS_SYM, TEMP_ELL, NBS_ELLIPSIS, PERIOD_ELLIPSIS, ELLIPSIS, NBSP, NBSPchar, aTAB As String
Public pBar As Progress_Bar
Public partialStatus, completeStatus As String
Public MyStoryNo As Integer
Public myRibbon As IRibbonUI
Public endCharCheck As Boolean
Public removeStyles(), replaceStyles(), skipStyles() As Variant

Function SetCharacters()

    MyStoryNo = 0

    'tab
    aTAB = vbTab
    'em dash
    EMDASH = ChrW(8212)
    'en dash
    ENDASH = ChrW(8211)
    'double prime
    DP = Chr(34)
    'single prime
    SP = Chr(39)
    'regular space
    aSPACE = Chr(32)
    'return character
    RTN = "[^013]"
    'return
    RTN2 = "^p"
    'temporary ellipsis substitution
    TEMP_ELL = "temp_ell"

    'Quote characters are different for Mac, so declare by OS for use in macros later
    #If Mac Then
        'single open quote
        SOQ = ChrW(8216)
        'single close quote
        SCQ = ChrW(8217)
        'double open quote
        DOQ = ChrW(8220)
        'double close quote
        DCQ = ChrW(8221)
        'ellipsis
        ELLIPSIS_SYM = ChrW(8230)
        'non-breaking space for search
        NBSP = "^s"
        'non-breaking space for typing
        NBSPchar = Chr(202)
        'ellipsis
        ELLIPSIS = "." & NBSPchar & "." & NBSPchar & "."
        NBS_ELLIPSIS = NBSPchar & ELLIPSIS
        PERIOD_ELLIPSIS = "." & NBSPchar & ELLIPSIS
        'non-breaking hyphen
        NBHYPH = ChrW(30)
        'optional hyphen
        OPTHYPH2 = ChrW(173)
        'optional hyphen
        OPTHYPH = ChrW(31)
    #Else
        'single open quote
        SOQ = Chr(145)
        'single close quote
        SCQ = Chr(146)
        'double open quote
        DOQ = Chr(147)
        'double close quote
        DCQ = Chr(148)
        'ellipsis symbol
        ELLIPSIS_SYM = Chr(133)
        'non-breaking space
        NBSP = Chr(160)
        'non-breaking space for typing
        NBSPchar = Chr(160)
        'ellipsis character
        ELLIPSIS = "." & NBSP & "." & NBSP & "."
        NBS_ELLIPSIS = NBSP & ELLIPSIS
        PERIOD_ELLIPSIS = "." & NBSP & ELLIPSIS
        'nonbreaking hyphen
        NBHYPH = ChrW(8209)
        'optional hyphen
        OPTHYPH = ChrW(173)
        'optional hyphen
        OPTHYPH2 = ChrW(31)
    #End If
    
End Function

Function DestroyCharacters()

    MyStoryNo = 0
    
    pStatus = vbNullString

    EMDASH = vbNullString
    ENDASH = vbNullString
    DP = vbNullString
    SP = vbNullString
    aSPACE = vbNullString
    RTN = vbNullString
    RTN2 = vbNullString
    TEMP_ELL = vbNullString
    SOQ = vbNullString
    SCQ = vbNullString
    DOQ = vbNullString
    DCQ = vbNullString
    ELLIPSIS_SYM = vbNullString
    NBS_ELLIPSIS = vbNullString
    PERIOD_ELLIPSIS = vbNullString
    NBSP = vbNullString
    NBSPchar = vbNullString
    ELLIPSIS = vbNullString
    aTAB = vbNullString
    NBHYPH = vbNullString
    OPTHYPH = vbNullString
    OPTHYPH2 = vbNullString
    
End Function
