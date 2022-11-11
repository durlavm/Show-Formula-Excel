Attribute VB_Name = "showFormulaModule"
Option Explicit

Public Function SHOW(ResultCell As Range, _
                        Optional SymbolOffset As Integer = 0, _
                        Optional AddSpaces As Boolean = False, _
                        Optional EqualSign As Boolean = True) As String
    'Returns the formula (equation) used with cell value or symbols
    'Argument #1 = ResultCell is the cell containing result calculation formula
    'Optional Argument #2 = SymbolOffset is # of columns to left/right of value cells
    'which contain the symbol for those values
    'Negative means to the left of cell values referenced
    'Positive means to the right of cell values referenced
    'Note that all value symbols should be same offset to left or right
    'since only one Optional SymbolOffset argument is passed
    'By default, SymbolOffset = 0, meaning this function will show equation with values
    'Optional Argument #3 = AddSpaces is Boolean to add spaces around operator signs
    'By default, the AddSpaces = False to not add any spaces
    'User should supply True as 3rd argument to add spaces for "cleaner" look
    'Operators impacted are +, -, *, /, ^, =
    'Optional Argument #4 = EqualSign is Boolean to show/hide equal sign of formula
    'By default, the EqualSign = True to display Equal (=) sign
    'User should supply False as 4th argument to remove the equal sign
    'Developed by Durlav Mudbhari
    'https://github.com/durlavm
    'https://www.linkedin.com/in/durlav1n1/

    'Application.Volatile
    'Volatile ensures this function gets recalculated when the worksheet updates
    
    Dim thisSht As Worksheet
    Set thisSht = ResultCell.Worksheet
    
    Dim cellEqn As String
    'Retrieve formula used as string
    cellEqn = ResultCell.Formula
    
    'Remove any absolute cell references from cell by removing $ signs
    cellEqn = Replace(cellEqn, "$", "")
    
    'If Optional Argument #4 = False then remove Equal (=) sign
    If EqualSign = False Then
    'Hide Equal Sign
        cellEqn = Replace(cellEqn, "=", "")
    End If
    
    'start with SHOW as blank string and build up the formula
    SHOW = ""
    
    Dim valChar As String 'each letter or symbol
    Dim alphaNum As String 'treat every alphabet as cell ref or keyword
    Dim cellRef As String 'store used cell ref to look up value
    
    Dim isAlphaNum As Boolean
    isAlphaNum = False  'assume non-alpha and non-numeric unless proved so
    Dim iChar As Integer
    iChar = 0
    Do While iChar < Len(cellEqn) 'check each character in the formula string
        alphaNum = ""
        iChar = iChar + 1
        valChar = Left(Mid(cellEqn, iChar), 1)  'extract char at the given position
        isAlphaNum = False
        
        Do While IsAlpha(valChar) = True Or IsNum(valChar) = True
            alphaNum = alphaNum + valChar   'start storing alphanumeric
            If IsNum(alphaNum) = True Then Exit Do
            'If alphaNum(alpha numeric) starts with num
            'then it cannot be a cell ref or keyword
            'it must be a numerica constant in equation,
            'so display it as is
            
            'if Do/Loop not exited then it must be alpha numeric
            isAlphaNum = True
            iChar = iChar + 1
            valChar = Left(Mid(cellEqn, iChar), 1)
        Loop
        
        'After Do/Loop exit, alphaNum should be concatenated
        If IsAlpha(alphaNum) = True Then
            'Check if alphaNum has no numbers
            'if no numbers then cant be a cell ref
            'so it must be a keyword of some sort
            iChar = iChar - 1
            
            If IsKeyword(alphaNum) = True Then
                'check if alphaNum is a std keyword with symbol
                'conver the alphanumeric keyword to symbol if possible
                alphaNum = keyWordSymb(alphaNum, SymbolOffset)
            End If
        
            'add keyword or symbol to existing equation
            SHOW = SHOW + alphaNum
            
        ElseIf isAlphaNum = True Then
            'If alphaNum is not purely alphabetical, meaning it has numbers
            'then it must be a cell ref
            'extract symbols or values from cell ref using SymbolOffset
            iChar = iChar - 1
            cellRef = thisSht.Range(alphaNum).Offset(0, SymbolOffset).Text
            'cellRef = thisSht.Range(alphaNum).Offset(0, SymbolOffset).Value2
            'Remove any equal sign present in the symbol cell
            cellRef = Replace(cellRef, "=", "")
            'Remove any spaces present in the symbol cell
            cellRef = Replace(cellRef, " ", "")
            
            'add cell value instead of alphaNum
            SHOW = SHOW + cellRef
        Else
            'If neither alpha or numeric then
            'it must be symbols in equation
            'then add it to equation as is
            SHOW = SHOW + valChar
        End If
    Loop
    'Equation should have been built one character at a time by end of the loop
    
    'If formula has () brackets, ex PI(), then remove ()
    If SHOW Like "*()*" = True Then
        SHOW = Replace(SHOW, "()", "")
    End If
    
    'If Optional Argument #3 = True then add spaces around operator symbols
    'If AddSpaces = True Then
        SHOW = Replace(SHOW, "+", " + ")
        SHOW = Replace(SHOW, "-", " - ")
        SHOW = Replace(SHOW, "*", " * ")
        SHOW = Replace(SHOW, "/", " \ ")
        SHOW = Replace(SHOW, "^", " ^ ")
        SHOW = Replace(SHOW, "=", "= ")
    'End If
    
End Function
 
Private Function IsAlpha(letter) As Boolean
    'Function to check whether a character is standard A thru Z alphabets
    'works for both lower and upper case alphabets
    IsAlpha = Len(letter) And Not letter Like "*[!A-Za-z]*"
End Function

Private Function IsNum(letter) As Boolean
    'Function to check whether a character is standard 0 thru 9 numerals
    IsNum = Len(letter) And Not letter Like "*[!0-9]*"
End Function

Private Function IsKeyword(word) As Boolean
    'Function to check if a supplied word is standard keyword
    'This check (If/then else) can be further expanded as needed
    IsKeyword = False
    If LCase(word) = "sqrt" Then
        IsKeyword = True
    ElseIf LCase(word) = "pi" Then
        IsKeyword = True
    ElseIf LCase(word) = "sum" Then
        IsKeyword = True
    End If
End Function

Private Function keyWordSymb(word, SymbolOffset) As String
    'Function to convert standard keywords to symbol as needed
    If LCase(word) = "sqrt" Then
        keyWordSymb = ChrW(8730)
    ElseIf LCase(word) = "sum" Then
        keyWordSymb = ChrW(931)
    ElseIf LCase(word) = "pi" And SymbolOffset = "0" Then
        keyWordSymb = "3.14"
    ElseIf LCase(word) = "pi" And SymbolOffset <> "0" Then
        keyWordSymb = ChrW(960)
    End If
End Function


