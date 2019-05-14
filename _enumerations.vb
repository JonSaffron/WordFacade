Public Enum WdAlertLevel
    wdAlertsNone = 0
    wdAlertsAll = -1
    wdAlertsMessageBox = -2
End Enum

Public Enum WdWindowState
    wdWindowStateNormal = 0
    wdWindowStateMaximize = 1
    wdWindowStateMinimize = 2
End Enum

Public Enum WdConstants
    wdTrue = -1
    wdFalse = 0
    wdToggle = 9999998
    wdUndefined = 9999999
End Enum

Public Enum WdUnderline
    wdUnderlineNone = 0
    wdUnderlineSingle = 1
    wdUnderlineWords = 2
    wdUnderlineDouble = 3
    wdUnderlineDotted = 4
    wdUnderlineThick = 6
    wdUnderlineDash = 7
    wdUnderlineDotDash = 9
    wdUnderlineDotDotDash = 10
    wdUnderlineWavy = 11
    wdUnderlineWavyHeavy = 27
    wdUnderlineDottedHeavy = 20
    wdUnderlineDashHeavy = 23
    wdUnderlineDotDashHeavy = 25
    wdUnderlineDotDotDashHeavy = 26
    wdUnderlineDashLong = 39
    wdUnderlineDashLongHeavy = 55
    wdUnderlineWavyDouble = 43
End Enum

Public Enum WdUnits
    wdCharacter = 1
    wdWord = 2
    wdSentence = 3
    wdParagraph = 4
    wdLine = 5
    wdStory = 6
    wdScreen = 7
    wdSection = 8
    wdColumn = 9
    wdRow = 10
    wdWindow = 11
    wdCell = 12
    wdCharacterFormatting = 13
    wdParagraphFormatting = 14
    wdTable = 15
    wdItem = 16
End Enum

Public Enum WdMovementType
    wdMove = 0
    wdExtend = 1
End Enum

Public Enum WdSeekView
    wdSeekMainDocument = 0
    wdSeekPrimaryHeader = 1
    wdSeekFirstPageHeader = 2
    wdSeekEvenPagesHeader = 3
    wdSeekPrimaryFooter = 4
    wdSeekFirstPageFooter = 5
    wdSeekEvenPagesFooter = 6
    wdSeekFootnotes = 7
    wdSeekEndnotes = 8
    wdSeekCurrentPageHeader = 9
    wdSeekCurrentPageFooter = 10
End Enum

Public Enum WdViewType
    wdNormalView = 1
    wdOutlineView = 2
    wdPrintView = 3
    wdPrintPreview = 4
    wdMasterView = 5
    wdWebView = 6
    wdReadingView = 7
End Enum

Public Enum WdBorderType
    wdBorderTop = -1
    wdBorderLeft = -2
    wdBorderBottom = -3
    wdBorderRight = -4
    wdBorderHorizontal = -5
    wdBorderVertical = -6
    wdBorderDiagonalDown = -7
    wdBorderDiagonalUp = -8
End Enum

Public Enum WdParagraphAlignment
    wdAlignParagraphCenter = 1
    wdAlignParagraphDistribute = 4
    wdAlignParagraphJustify = 3
    wdAlignParagraphJustifyHi = 7
    wdAlignParagraphJustifyLow = 8
    wdAlignParagraphJustifyMed = 5
    wdAlignParagraphLeft = 0
    wdAlignParagraphRight = 2
End Enum

Public Enum WdColor
    wdColorBlack = 0
    wdColorRed = &HFF
    wdColorGray05 = &HF3F3F3
End Enum

Public Enum WdDefaultTableBehavior
    wdWord8TableBehavior = 0
    wdWord9TableBehavior = 1
End Enum

Public Enum WdAutoFitBehavior
    wdAutoFitContent = 1
    wdAutoFitFixed = 0
    wdAutoFitWindow = 2
End Enum

Public Enum WdOrientation
    wdOrientLandscape = 1
    wdOrientPortrait = 0
End Enum

Public Enum WdBreakType
    wdPageBreak = 7
End Enum

Public Enum WdTabAlignment
    wdAlignTabBar = 4
    wdAlignTabCenter = 1
    wdAlignTabDecimal = 3
    wdAlignTabLeft = 0
    wdAlignTabList = 6
    wdAlignTabRight = 2
End Enum

Public Enum WdTabLeader
    wdTabLeaderDashes = 2
    wdTabLeaderDots = 1
    wdTabLeaderHeavy = 4
    wdTabLeaderLines = 3
    wdTabLeaderMiddleDot = 5
    wdTabLeaderSpaces = 0
End Enum

Public Enum WdPreferredWidthType
    wdPreferredWidthAuto = 1
    wdPreferredWidthPercent = 2
    wdPreferredWidthPoints = 3
End Enum

Public Enum WdCellVerticalAlignment
    wdCellAlignVerticalBottom = 3
    wdCellAlignVerticalCenter = 1
    wdCellAlignVerticalTop = 0
End Enum

Public Enum WdTextureIndex
    wdTextureNone = 0
    wdTextureSolid = 1000
End Enum

Public Enum WdColorIndex
    wdAuto = 0
    wdBlack = 1
End Enum

Public Enum WdLineStyle
    wdLineStyleNone = 0
    wdLineStyleSingle = 1
End Enum

Public Enum WdLineWidth
    wdLineWidth050pt = 4
End Enum

Public Enum WdSaveOptions
    wdDoNotSaveChanges = 0
    wdSaveChanges = -1
    wdPromptToSaveChanges = -2
End Enum

Public Enum MsoDocProperties
    msoPropertyTypeNumber = 1
    msoPropertyTypeBoolean = 2
    msoPropertyTypeDate = 3
    msoPropertyTypeString = 4
    msoPropertyTypeFloat = 5
End Enum

Public Enum WdSectionStart
    wdSectionContinuous = 0
    wdSectionNewColumn = 1
    wdSectionNewPage = 2
    wdSectionEvenPage = 3
    wdSectionOddPage = 4
End Enum

Public Enum WdHeaderFooterIndex
    wdHeaderFooterPrimary = 1
    wdHeaderFooterFirstPage = 2
    wdHeaderFooterEvenPages = 3
End Enum

Public Enum WdSaveFormat
    wdFormatDocument = 0                    ' .doc format

    wdFormatDocumentDefault = 16            ' .docx format under word 2007
End Enum

Public Enum WdTableFieldSeparator
    wdSeparateByParagraphs = 0
    wdSeparateByTabs = 1
    wdSeparateByCommas = 2
    wdSeparateByDefaultListSeparator = 3
End Enum

Public Enum WdCollapseDirection
    wdCollapseEnd = 0
    wdCollapseStart = 1
End Enum

Public Enum WdFindWrap
    wdFindAsk = 2
    wdFindContinue = 1
    wdFindStop = 0
End Enum

Public Enum WdReplace
    wdReplaceAll = 2
    wdReplaceNone = 0
    wdReplaceOne = 1
End Enum

Public Enum WdBuiltInStyle
    wdStyleNormal = -1
    wdStyleHeading1 = -2
    wdStyleHeading2 = -3
    ' plenty more of these that could be added
End Enum

Public Enum WdAnimation
    wdAnimationNone = 0
End Enum

Public Enum WdLineSpacing
    wdLineSpaceSingle = 0
End Enum

Public Enum WdOutlineLevel
    wdOutlineLevel1 = 1
    wdOutlineLevel2 = 2
End Enum
