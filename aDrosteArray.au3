#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=aDrosteArray.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Droste array functions
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Bert Kerkhof 21 april 2018
#AutoIt3Wrapper_Res_SaveSource=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 3 -w 4 -w 5
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 2 /reel
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv /rm
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include-once
#include <StringConstants.au3> ; Delivered with AutoIT
#include <FileConstants.au3> ; Delivered with AutoIT

; Written with AutoIT v3.3.14.2 interpreter/compiler

; aDrosteArray module : basic array-in-array functions ==========================
;    These differ from zero based array routines delivered with many traditional
;    computer compilers. They also differ with the AutoIt _array package:
;
; A. The zero place in aDrosteArray is reserved for the total number of
;    elements. This field is kept up-to-date by the DrosteArray routines.
;    So whether the user deletes an element with aDelete() or inserts a element
;    with aInsert(), the zero-element always contains the highest index number.
;
;    With a DrosteArray, the user does not have to calculate the number of user
;    elements with +1 nor the end-argument in a for-next loop with -1 anymore.
;    This reduces the number of runtime out-of-bound messages. The total number
;    of user-elements for every array is close at hand in the aArray[0] element.
;    With less worry about these details you can concentrate on design and
;    efficiency of the program.
;
; B. Droste multi-dimensional arrays are more versatile than square 2D arrays.
;    The option of different row sizes is more adapt to many logistics which
;    software designers face. Adding elements and specific dimensions reflect
;    real life situations which asks for detail.
;
; C. Many available routines operate on a single dimension array only. The
;    array-in-array design makes it easy to retrieve a single array and restore
;    it afterwards in multidimensional space. This reduces the need to do
;    double coding in routines for both one- and two dimensional data input.
;
; D. With the aAdd and build functions, your program will never run stuck on
;    preprogrammed dimensional limitations or square sizes. Droste array's add
;    a level of abstraction. Focus on basics avoids a waste of subscript
;    numbers. Using lists in stead of single values will propel productivity.
;    Dorste makes it work almost as pleasant as the use of strings.
;
;    The module is named after a dutch chocolate maker. For the recursive visual
;    effect, see:  https://en.wikipedia.org/wiki/Droste_effect
;
;       Bert Kerkhof
;
;
; Contents:
;    Create array's and add:
;       + aNewArray       + Creates a Droste array (one or two dimensions)
;       + aAdd            + Add value (single element or array) to an array
;       + aD              + Add a row of multiple values to a 2dim array
;       + aConcat         + Initialize an array, quickly assign values
;       + aRecite         + Place string values separated by '|' in array
;    Display array's:
;       + rRecite         + Place array values in a single line of text
;    Matrix operations on rows and columns:
;       + aLeft           + Copy the first items from an array
;       + aRight          + Copy the last items from an array
;       + aMid            + Copy items from an array
;       + aRound          + Converts an array of values to array of numbers
;       + aDec            + Convert hex values in a string to an array of numbers
;       + aColumn         + Copy single column from two-dimension array.
;       + aCompare        + Compare contents of two multi-dim arrays
;    Operations on matrix elements:
;       + GetA            + Retrieve value from a 2dim array
;       + SetA            + Sets value in a 2dim array
;       + IncrA           + Increments value in a 2dim array
;       + aMax            + Returns the highest value in an array
;       + aMin            + Returns the lowest value in an array
;       + LastIndex       + Returns highest index number (length) of array
;    Search and find:
;       + aSearch         + Search element or row in array
;       + aFindAll        + Collect multiple index positions in an array
;       + aStringFindAll  + Collect multiple string index positions in an array
;    Array modifications:
;       + aDelete         + Delete element or row in an array
;       + aInsert         + Insert element or row in an array
;       + aFlat           + Collect multiple values and arrays in a single array
;       + aFill           + Fill array with value
;    Sort:
;       + aCombSort       + Sorts a two-dimensional array
;       + aSimpleCombSort + Sorts a one-dimensional array
;    Array conversion:
;       + aDroste         + Convert zero-based array to type Droste
;       + StringUtfArray  + Converts one-dim array of utf16 values to string
;       + ArrayUtfString  + Converts string to one-dim array of utf16 values
;       + StringFromArray + Converts (multi dim) array to a single string
;       + ArrayFromString + Restores a (multi dim) array from a single string
;    Array store and restore:
;       + WriteArray      + Write a (multi dim) array to file
;       + ReadArray       + Read a (multi dim) array from file

; Author ....: Bert Kerkhof ( kerkhof.bert@gmail.com )
; Tested with: AutoIt version 3.3142 and win10 / win7

#Region ; Create and build ======================================================

Func aNewArray($nDim1 = 0, $nDim2 = 0) ; Create a new array
  ; Syntax.........: aNewArray(Dim1 [, Dim2])
  ; Parameters ....: - Dim1 (optional) last index (length) in the first dimension
  ;                    If omitted, array will have 0 elements
  ;                  - Dim2 (optional) of the second dimension.
  ;                    If omitted, a one-dimensional array is created
  ; Return values .: Array with the specified dimension(s)
  ; See also ......: aAdd aConcat aRecite

  Local $aR1[1] = [0], $aR2[1] = [0]
  ReDim $aR1[$nDim1 + 1]
  $aR1[0] = $nDim1
  If @NumParams = 2 Then
    ReDim $aR2[$nDim2 + 1]
    $aR2[0] = $nDim2
  Else
    $aR2 = 0
  EndIf
  aFill($aR1, $aR2)
  Return $aR1
EndFunc   ;==>aNewArray

Func aAdd(ByRef $aArray, $Row)
  ; Add value(s) at the end of given array
  ; Syntax.........: aAdd(aArray, nValue) adds one value
  ;                : aAdd(aArray1, aArray2) adds a row to the last $aArray1 element
  ;                : If the added value is an array, aArray1 is two-dimensional
  ; Parameters ....: aArray  - Array to modify
  ;                  Value   - Value to add
  ; Returns...... .: The new length of aArray
  ; Author ........: Bert Kerkhof
  ; Comment .......: Because the AutoIT implementation of ReDim is fast, aAdd is an
  ;                  efficient way to build your custom sized array's from scratch

  If _check1($aArray) Then Return _errmess("aAdd", $aArray[0])
  Local $I = $aArray[0] + 1
  ReDim $aArray[$I + 1]
  $aArray[$I] = $Row
  $aArray[0] = $I
  Return $I
EndFunc   ;==>aAdd

Func aConcat($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0)
  ; Description ...: Concatenates up to 10 values: numbers / strings / arrays or a mixture
  ;                : If a value is an array, the resulting array is multi-dimensional
  ; Syntax.........: Array($v1 [,$v2 [,.. [, $v10 ]]])
  ; Parameters ....: $v1  - First element of the array
  ;                  $v2  - [optional] Second element of the array
  ;                  ...
  ;                  $v10 - [optional] Tenth element of the array
  ; Returns .......: Array with values, the zero element contains the number of values
  ; See also ......: aRecite aFlat aNewArray aAdd

  Local $aArray = [@NumParams, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J]
  ReDim $aArray[$aArray[0] + 1]
  Return $aArray
EndFunc   ;==>aConcat

Func aD($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0)
  ; Description ...: Builds two dimensional array. Each row contains up to 10 values:
  ;                  numbers / strings / arrays or mixture. When called with zero
  ;                  parameters, the build-up is returned and re-initialized.
  ; Returns .......: Array with a two dimensional array of values
  ; Author ........: Bert Kerkhof
  Local $nParam = @NumParams
  Local Static $aSave = aNewArray(0)
  If $nParam Then
    Local $aB = [$nParam, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J]
    aAdd($aSave, aLeft($aB, $nParam))
  Else
    Local $aR = $aSave
    $aSave = aNewArray(0)
    Return $aR
  EndIf
EndFunc   ;==>aD

Func aRecite(Const $rString, $Sep = '|')
  ; Convert a recite of values to array
  ; Also see ......: rRecite aDec
  Return StringSplit($rString, $Sep) ; ® Split on vertical bar (not on comma)
  ; The result of StringSplit is copied to enable possible writing to its elements:
  ; Local $aArray = aNewArray($aSplit[0])
  ; For $I = 1 To $aSplit[0]
  ;   $aArray[$I] = $aSplit[$I]
  ; Next
  ; Return $aArray
EndFunc   ;==>aRecite

#EndRegion ; Create and build ======================================================

#Region ; Display ===============================================================

Func rRecite(Const $aArray, $sSeparator = '|', $sPrefix = '', $sPostfix = '')
  ; Readable array values on a single text line
  ; Syntax.........: rRecite(aArray, Separator, QuoteFlag, Fetch)
  ; Parameters ....: aArray      - Series of values to put readable on a line
  ;                  sSeparator  - [Optional] String that separates values,
  ;                                default : vertical bar
  ;                  sPrefix     - [Optional] String will prefix each element
  ;                  sPostFix    - [Optional] String will postfix each element
  ; Returns        : Readable string

  If _check1($aArray) Then Return _errmess("rRecite", '')
  Local $I, $sResult = ''
  For $I = 1 To $aArray[0]
    $sResult &= $sPrefix & $aArray[$I] & $sPostfix
    If $I < $aArray[0] Then $sResult &= $sSeparator
  Next
  Return $sResult ; ® readable result
EndFunc   ;==>rRecite

#EndRegion ; Display ===============================================================

#Region ; Matrix operations on rows and columns =================================

Func aColumn($aaArray, $iColumn, $iStart = 1, $iEnd = 0)
  ; Create single dim array from column in a multi-dim array
  If _check4($aaArray, $iStart, $iEnd, $iColumn) Then Return _errmess("aColumn", aNewArray(0))
  Local $N = $iEnd - $iStart + 1, $nOffset = $iStart - 1
  Local $I, $aR = aNewArray($N)
  For $I = 1 To $N
    Local $aArray = $aaArray[$I + $nOffset]
    $aR[$I] = $aArray[$iColumn]
  Next
  Return $aR
EndFunc   ;==>aColumn

Func aLeft($aArray, $N = 1)
  ; Returns the first $N elements of an array:
  If _check1($aArray) Then Return _errmess("aLeft", aNewArray(0))
  $aArray[0] = Max(0, Min($aArray[0], $N))
  ReDim $aArray[$aArray[0] + 1]
  Return $aArray
EndFunc   ;==>aLeft

Func aRight($aIn, $N = 1)
  ; Returns the last $N elements of an array:
  If _check1($aIn) Then Return _errmess("aRight", aNewArray(0))
  $N = Max(0, Min($N, $aIn[0]))
  Local $I, $aOut = aNewArray($N), $nOffset = $aIn[0] - $N
  For $I = 1 To $N
    $aOut[$I] = $aIn[$I + $nOffset]
  Next
  Return $aOut
EndFunc   ;==>aRight

Func aMid($aIn, $iPos, $N = 999999999999)
  ; Returns $N elements of an array, starting at $iPos:
  If _check2($aIn, $iPos) Then Return _errmess("aMid", aNewArray(0))
  $N = Max(0, Min($N, $aIn[0] - $iPos + 1))
  Local $I, $aOut = aNewArray($N)
  For $I = 1 To $N
    $aOut[$I] = $aIn[$I + $iPos - 1]
  Next
  Return $aOut
EndFunc   ;==>aMid

Func aRound($aArray, $nDecimal = 0)
  ; Converts an array of values to type number
  If _check1($aArray) Then Return _errmess("aRound", aNewArray(0))
  Local $I
  For $I = 1 To $aArray[0]
    $aArray[$I] = Round($aArray[$I], $nDecimal)
  Next
  Return $aArray
EndFunc   ;==>aRound

Func aDec(Const $sString, $sSep = '|')
  ; Convert hexadecimal values held in a string to array of numbers
  ; Related .......: Dec
  Local $aR = StringSplit($sString, $sSep) ; Split on vertical bar char
  For $I = 1 To $aR[0]
    $aR[$I] = Dec($aR[$I])
  Next
  Return $aR
EndFunc   ;==>aDec

Func LastIndex(Const $aArray)
  ; Returns the highest index number of array
  ; Syntax.........: LastIndex(aArray)
  ; Return values .: Index number
  If _check1($aArray) Then Return _errmess("LastIndex", 0)
  Return $aArray[0]
EndFunc   ;==>LastIndex

Func aMax(Const $aArray, $iStart = 1, $iEnd = 0)
  ; Returns the highest value held in the numeric array
  ; Syntax.........: aMax($aArray[, $iStart = 1[, $iEnd = 0]])
  ; Parameters ....: $aArray  - Array to search
  ;                  $iStart  - [optional] Index of array to start searching
  ;                  $iEnd    - [optional] Index of array to stop searching
  ; Return values .: The maximum value
  ; Also see ......: aMin, Max

  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("aMax", 0)
  Local $I, $iMaxIndex = $iStart
  For $I = $iStart + 1 To $iEnd
    If $aArray[$iMaxIndex] < $aArray[$I] Then $iMaxIndex = $I
  Next
  Return $aArray[$iMaxIndex]
EndFunc   ;==>aMax

Func aMin(Const $aArray, $iStart = 1, $iEnd = 0)
  ; Returns the lowest value held in the numeric array.
  ; Syntax.........: aMin($aArray[, $iStart = 1[, $iEnd = 0]])
  ; Parameters ....: $aArray  - Array to search
  ;                  $iStart  - [optional] Index of array to start searching
  ;                  $iEnd    - [optional] Index of array to stop searching
  ; Return values .: The minimum value
  ; Also see ......: aMax, Min

  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("aMin", 0)
  Local $I, $iMaxIndex = $iStart
  For $I = $iStart + 1 To $iEnd
    If $aArray[$iMaxIndex] > $aArray[$I] Then $iMaxIndex = $I
  Next
  Return $aArray[$iMaxIndex]
EndFunc   ;==>aMin

#EndRegion ; Matrix operations on rows and columns =================================

#Region ; Operations on matrix elements =========================================

Func GetA(Const $aR, $iP = 1)
  ; Get value from a (2dim) array
  ; Comment .......: Useful for Droste 2dim array
  ; Also see ......: aSet, aIncr
  ; Author ........: Thanks Wouter Bos
  If _check1($aR) Or $iP < 0 Or $iP > $aR[0] Then Return _errmess("GetA", '')
  Return $aR[$iP]
EndFunc   ;==>GetA

Func SetA(ByRef $aR, $iP = 1, $Val = 0)
  ; Set value in a (2dim) array
  ; Comment .......: Useful for Droste 2dim array
  ; Also see ......: aGet, aIncr
  If _check1($aR) Or $iP < 0 Or $iP > $aR[0] Then Return _errmess("SetA", '')
  $aR[$iP] = $Val
  Return $aR[$iP]
EndFunc   ;==>SetA

Func IncrA(ByRef $aR, $iP = 1, $N = 1)
  ; Increments value in a (2dim) array with +/- $N
  ; Comment .......: Useful for Droste 2dim array
  ; Also see ......: aGet, aSet
  If _check1($aR) Or $iP < 0 Or $iP > $aR[0] Then Return _errmess("IncrA", '')
  $aR[$iP] += $N
  Return $aR[$iP]
EndFunc   ;==>IncrA

Func Max($N1, $N2) ; Compare values and return the highest one
  Local $R = $N1
  If $N2 > $N1 Then $R = $N2
  Return $R
EndFunc   ;==>Max

Func Min($N1, $N2) ; Compare values and return the lowest one
  Local $R = $N1
  If $N2 < $N1 Then $R = $N2
  Return $R
EndFunc   ;==>Min

#EndRegion ; Operations on matrix elements =========================================

#Region ; Search and find =======================================================

Func aSearch($aArray, $vValue, $iColumn = 0, $iStart = 1, $iEnd = 0)
  ; Search element in 1dim / 2dim array
  ; Syntax.........: aSearch(aArray, vValue, dummy, $iStart=1, $iEnd=0)
  ; Parameters ....: $aArray  - The array to search
  ;                  $iColumn - Zero for 1dim array
  ;                             2dim: column to search
  ;                  $vValue  - What to search for in $aArray
  ;                  $iStart  - [optional] Index of array to start searching
  ;                  $iEnd    - [optional] Index of array to stop searching
  ; Return values .: Found : Index number of the string found
  ;                  Not found : 0
  ; Remark ........: Case insensitive
  ; See also ......: aFindAll aStringFindAll StringInStr

  If _check4($aArray, $iStart, $iEnd, $iColumn) Then Return _errmess("aSearch", 0)
  Local $I, $nOffset = 0
  If $iColumn Then
    $aArray = aColumn($aArray, $iColumn, $iStart, $iEnd)
    $nOffset = $iStart - 1
    $iStart = 1
    $iEnd = $aArray[0]
  EndIf
  For $I = $iStart To $iEnd ; Case sesnsitive
    If $aArray[$I] = $vValue Then Return $I + $nOffset
  Next
  Return SetError(6, 0, 0)
EndFunc   ;==>aSearch

Func aFindAll(Const $aArray, $Value, $iStart = 1, $iEnd = 0)
  ; Collect multiple occurrences in an array
  ; Remark ........: Case insensitive
  ; Related .......: aStringFindAll aSearch

  Local $I, $aR = aNewArray(0)
  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("aFindAll", $aR)
  For $I = $iStart To $iEnd
    If $aArray[$I] = $Value Then aAdd($aR, $I)
  Next
  Return $aR
EndFunc   ;==>aFindAll

Func aStringFindAll($sString, $sSearch, $iStart = 1, $iEnd = 0)
  ; Collect multiple occurrences in an array
  ; Remark ........: Case insensitive
  ; Related .......: aSearch, aFindAll

  If $iStart < 1 Or $iStart > StringLen($sString) Then $iStart = 1
  If $iEnd = 0 Then $iEnd = StringLen($sString)
  Local $nOffset, $aR = aNewArray(0)
  While True
    $nOffset = StringInStr($sString, $sSearch, $STR_NOCASESENSE, 1, $iStart, $iEnd - $iStart + 1)
    If $nOffset = 0 Then ExitLoop
    aAdd($aR, $nOffset)
    $iStart = $nOffset + 1
  WEnd
  Return $aR
EndFunc   ;==>aStringFindAll

#EndRegion ; Search and find =======================================================

#Region ; Array modification ====================================================

Func aDelete(ByRef $aArray, $iPos = 1, $N = 1)
  ; Delete one or more elements in an array
  ; Syntax.........: aDelete(aArray, iElement, N)
  ; Parameters ....: aArray  - Array to modify
  ;                  iPos    - Element to delete
  ;                            By default the first
  ;                  N       - Number of elements to delete
  ;                            By default one
  ; Returns .......: Index of last item, equals the new length
  ; Comment .......: Deleting elements one by one is a slow process.
  ;                  Try to delete multiple elements at once.

  If _check2($aArray, $iPos) Then Return _errmess("aDelete", $aArray)
  $N = Min($N, $aArray[0] - $iPos + 1)
  If $iPos < 1 Or $N < 1 Then Return
  Local $J, $I = $aArray[0] - $N
  For $J = $iPos To $I
    $aArray[$J] = $aArray[$J + $N]
  Next
  ReDim $aArray[$I + 1]
  $aArray[0] = $I ; Update number of last index in element zero
  Return $aArray[0]
EndFunc   ;==>aDelete

Func aInsert(ByRef $aArray, $iPos = 1, $Value = 0)
  ; Insert one or an array of elements in an array
  ; Syntax.....: aInsert(aArray, iElement, nElement, Value)
  ; Parameters : aArra - Array to modify
  ;              iPos  - Element number to insert row(s)
  ;                      by default insertion is in the first element
  ;                      Insertion possible up to LastIndex($aArray) +1
  ;              Value - Value to insert.
  ;                      If Value is an array, all its elements are inserted.
  ;                      This causes the length of the receiving array to
  ;                      increase with the number of elements.
  ;                      If Value is a single element, only this one is
  ;                      is inserted.
  ; Returns ...: The new array length.
  ;              The resulting array is passed by reference.
  ; Example ...: aFlat()
  ; Comment ...: Inserting elements one by one is a slow process.
  ;              Try to insert multiple elements at once.

  If _check1($aArray) Then Return _errmess("aInsert", $aArray)
  If $iPos < 1 Then $iPos = 1
  If $iPos > $aArray[0] + 1 Then $iPos = $aArray[0] + 1

  Local $J, $vPos
  Local $nSize = IsArray($Value) ? $Value[0] : 1
  $aArray[0] += $nSize ; Increase length of the array
  ReDim $aArray[$aArray[0] + 1]
  For $J = $aArray[0] To $iPos + $nSize Step -1
    $aArray[$J] = $aArray[$J - $nSize]
  Next
  If IsArray($Value) Then
    $vPos = 1
    For $J = $iPos To $iPos + $nSize - 1
      $aArray[$J] = $Value[$vPos]
      $vPos += 1
    Next
  Else
    $aArray[$iPos] = $Value
  EndIf
  Return $aArray[0]
EndFunc   ;==>aInsert

Func aFlat($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0)
  ; Description ...: Serially concatenates up to 10 array's / numbers / strings
  ; Syntax.........: aSerialConcat(aArray1, aArray1, Value .. $aArrayN)
  ; Returns .......: One-dimensional array of values
  ; See also ......: aRecite aConcat aNewArray aAdd
  ; Author ........: Bert Kerkhof

  Local $L, $aReturn = aNewArray(0)
  Local $aArrayB = [@NumParams, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J]
  For $L = 1 To $aArrayB[0]
    aInsert($aReturn, $aReturn[0] + 1, $aArrayB[$L]) ; Insertion point is Lastindex($aArray)+1
  Next
  Return $aReturn
EndFunc   ;==>aFlat

Func aFill(ByRef $aArray, Const $vValue = 0)
  ; Fill an array with value
  ; Syntax.........: Afill(aArray [,vValue])
  ; Parameters ....: aArray  - Array to be modified
  ;                  aValue  - [optional] value to fill. If omitted, zero is assumed
  ; Comment .......: All elements are filled, even if aArray has multiple dimensions
  ; Return values .: Returns number of updated elements
  If _check1($aArray) Then Return _errmess("aFill", 0)
  Local $I, $Row, $N = 0
  For $I = 1 To $aArray[0]
    $Row = $aArray[$I]
    If IsArray($Row) Then
      $N += aFill($Row, $vValue) ; Recurse
    Else
      $aArray[$I] = $vValue
      $N += 1
    EndIf
  Next
  Return $N
EndFunc   ;==>aFill

#EndRegion ; Array modification ====================================================

#Region ; Array conversion ======================================================

Func aDroste(Const $aArray) ; Convert 1dim array to type Droste
  ; Syntax.........: aDroste(aArray)
  ; Parameters ....: aArray  - Array to be modified
  ; Comment .......: A first element is added with the length of the array
  ; Return values .: The Drosted array
  ; Remarks .......: Execution time 135 milliseconds on quad computer for
  ;                  100.000 elements array

  If Not IsArray($aArray) Then Return _errmess("aDroste", aNewArray(0))
  Local $I, $NewLen = UBound($aArray)
  Local $aReturn[$NewLen + 1]
  For $I = 1 To $NewLen
    $aReturn[$I] = $aArray[$I - 1]
  Next
  $aReturn[0] = $NewLen
  Return $aReturn ;
EndFunc   ;==>aDroste

Func StringUtfArray($aArray, $iStart = 1, $iEnd = 0)
  ; Convert an array of utf16 numbers into a string
  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("StringFromArray", '')
  ; Due to an inconvenience, $iEnd has to surpass uBound:
  Return StringFromASCIIArray($aArray, $iStart, $iEnd + 1, 0)
EndFunc   ;==>StringUtfArray

Func ArrayUtfString($sString, $iStart = 1, $iEnd = 0)
  ; Convert a string to an array of utf16 numbers
  If $iStart < 1 Then $iStart = 1
  If $iEnd > StringLen($sString) Then $iEnd = StringLen($sString)
  ; Due to an inconvenience, $iStart=0 points to first char in string:
  Return aDroste(StringToASCIIArray($sString, $iStart - 1, $iEnd))
EndFunc   ;==>ArrayUtfString

#EndRegion ; Array conversion ======================================================

#Region ; Sort ==================================================================

Func NumberCompare($N1, $N2)
  ; helper function for aCombSort :
  If $N1 > $N2 Then Return 1
  If $N1 < $N2 Then Return -1
  Return 0
EndFunc   ;==>NumberCompare

; Sort a two-dimensional array :
Func aCombSort(ByRef $aArray, $iColumn = 1, $Descending = True, $Numeric = False, $iStart = 1, $iEnd = 0)
  ; Syntax.........: aCombSort(Array, Column, Orderflag, NumericFlag, $iStart, $iEnd)
  ; Parameters ....: + aArray       : Two-dimensional array
  ;                  + iColumn      : Column number of the array that contains the sort key
  ;                  + Descending   : (optional) Sort order flag.
  ;                                   Default: [True] low values sorted first
  ;                                   [False] low values sorted last
  ;                  + Numeric      : (optional) Default: [False] sort key is string
  ;                  + iStart, iEnd : (optional) Range of elements to be sorted
  ; Return value...: Number of passes.
  ;                  The sorted array is passed by reference
  ; Author ........: Bert Kerkhof

  If _check4($aArray, $iStart, $iEnd, $iColumn) Then Return _errmess("aCombSort", 0)
  If $iStart > $iEnd Then Return 0 ; Zero elements in array

  Local $Gap = $iEnd - $iStart ; Initialize gap size
  Local $nPass = 0, $R1, $R2
  Local $CompareLogic = $Descending ? -1 : 1
  Do
    $nPass += 1
    $Gap = Max(1, Floor($Gap / 1.3)) ; Apply the gap shrink factor
    Local $Count = $iStart, $Virgin = ($Gap = 1)
    If $Numeric Then ; Sort key is type numeric
      While $Count + $Gap <= $iEnd ; Single 'comb' over the input list
        $R1 = $aArray[$Count]
        $R2 = $aArray[$Count + $Gap]
        If NumberCompare($R1[$iColumn], $R2[$iColumn]) = $CompareLogic Then
          $aArray[$Count] = $R2
          $aArray[$Count + $Gap] = $R1
          $Virgin = False
        EndIf
        $Count += 1
      WEnd
    Else ; Sort key is type string
      While $Count + $Gap <= $iEnd ; Single 'comb' over the input list
        $R1 = $aArray[$Count]
        $R2 = $aArray[$Count + $Gap]
        If StringCompare($R1[$iColumn], $R2[$iColumn]) = $CompareLogic Then
          $aArray[$Count] = $R2
          $aArray[$Count + $Gap] = $R1
          $Virgin = False
        EndIf
        $Count += 1
      WEnd
    EndIf
  Until $Virgin
  Return $nPass
EndFunc   ;==>aCombSort

Func aSimpleCombSort(ByRef $aArray, $Descending = True, $Numeric = False, $iStart = 1, $iEnd = 0)
  ; Perform aCombSort on a one-dimensional array
  ;   + Descending : (optional) Sort order flag.
  ;                  Default: [True] low values sorted first
  ;                  [False] low values sorted last
  ;   + Numeric    : (optional) Default: [False] sort key is string
  ; Return value...: Number of passes.
  ;                  The sorted array is passed by reference

  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("aSimpleCombSort", 0)
  Local $I, $aRelem
  For $I = $iStart To $iEnd ; Pack :
    $aArray[$I] = aConcat($aArray[$I])
  Next
  Local $nPass = aCombSort($aArray, 1, $Descending, $Numeric, $iStart, $iEnd)
  For $I = $iStart To $iEnd ; UnPack :
    $aRelem = $aArray[$I]
    $aArray[$I] = $aRelem[1]
  Next
  Return $nPass
EndFunc   ;==>aSimpleCombSort

#EndRegion ; Sort ==================================================================

#Region ; Array <=> String and file  ===========================================

Func StringFromArray($aR, $S = ChrW(9))
  ; Restore (multi dim) array from a single string
  ; Parameters ..: $aR     : Previous encoded data, see StringFromArray()
  ;                $S      : Unique separator char
  ;                          Choose carefully in case of binary data
  ;                          Default: TAB
  ; Returns .....: $String : Result
  ; Author ......: Bert Kerkhof

  If _check1($aR) Then Return _errmess("StringFromArray", False)
  If StringLen($S) = 1 Then $S &= '0'
  Local $SN = StringLeft($S, 1) & ChrW(AscW(StringRight($S, 1)) + 1) ; Marker
  Local $I, $String = ''
  For $I = 1 To $aR[0]
    If IsArray($aR[$I]) Then
      $String &= $aR[0] = 1 ? $SN & StringFromArray($aR[$I], $SN) : StringFromArray($aR[$I], $SN)
    Else
      $String &= $aR[$I]
    EndIf
    If $I < $aR[0] Then $String &= $SN
  Next
  Return $String
EndFunc   ;==>StringFromArray

Func ArrayFromString($String, $S = ChrW(9))
  ; Store the contents of (multi dim) array in a single string
  ; Parameters ..: $String : Previous encoded data, see StringFromArray()
  ;                $S      : Unique separator char
  ;                          Use the same as in StringFromArray()
  ;                          Default: TAB
  ; Returns .....: $aR     : Result array
  ; Author ......: Bert Kerkhof

  If StringLen($S) = 1 Then $S &= '1'
  Local $SN = StringLeft($S, 1) & ChrW(AscW(StringRight($S, 1)) + 1) ; Marker
  If StringLeft($String, 2) = $S Then Return aConcat(ArrayFromString(StringMid($String, 3), $SN))
  Local $aSp = StringSplit($String, $S, $STR_ENTIRESPLIT)
  Local $I, $aR = aNewArray($aSp[0])
  For $I = 1 To $aSp[0] ; Recurse:
    $aR[$I] = StringInStr($aSp[$I], $SN) ? ArrayFromString($aSp[$I], $SN) : $aSp[$I]
  Next
  Return $aR
EndFunc   ;==>ArrayFromString

Func CompareArray(Const $aArray1, Const $aArray2)
  ; Compare the content of two (multi dim) Droste arrays
  ; Syntax ........: aCompare($aArray1, $aArray2)
  ; Parameters ....: $aArray1 and $aArray2
  ; Return value ..: True if equal, False if unequal
  ; Author ........: Thanks Kim Putters
  ; Remark ........: Case insensitive

  If _check1($aArray1) Then Return _errmess("aCompare", False)
  If _check1($aArray2) Then Return _errmess("aCompare", False)
  Return StringFromArray($aArray1) = StringFromArray($aArray2)
EndFunc   ;==>CompareArray

Func WriteArray($Array, $FileName)
  ; Write (multi dim) array to file
  ; Parameters ....: $Array    The (multi dim) array to be written
  ;                  $FileName Full path
  ; Returns .......: @error flag is set to 1 in case of any failure
  Local $S = StringFromArray($Array)
  If @error Then Return SetError(1)
  Local $H = FileOpen($FileName, $FO_UTF8 + $FO_CREATEPATH + $FO_OVERWRITE)
  If @error Then Return SetError(1)
  If FileWrite($H, $S) = 0 Then Return SetError(1)
  FileClose($H)
EndFunc   ;==>WriteArray

Func ReadArray($FileName)
  ; Read (multi dim) array from file
  ; Parameter ......: $FileName Full path
  ; Returns ........: A (multi dim) array
  ;                 : @error flag is set to 1 in case of any failure
  Local $H = FileOpen($FileName, $FO_READ)
  If @error Then Return SetError(1)
  Local $S = FileRead($H)
  If @error Then Return SetError(1)
  FileClose($H)
  Return ArrayFromString($S)
EndFunc   ;==>ReadArray

#EndRegion ; Array <=> String and file  ===========================================

#Region ; Safeguard boundaries and check array =================================

Func _errmess($sOrigin, $Rvalue)
  If @Compiled Then Return SetError(1, 0, $Rvalue)
  MsgBox(64, 'aDrosteArray', 'error with array' & @CRLF & 'in function ' & $sOrigin)
  Exit
  Return SetError(1, 0, $Rvalue)
EndFunc   ;==>_errmess

Func _check1($aArray)
  ; Check validity of $aArray :
  If Not IsArray($aArray) Then Return True
  ; Un-comment this if you are unfamiliar with Droste array's :
  ; If UBound($aArray) <> $aArray[0] + 1 Then Return True
  Return False
EndFunc   ;==>_check1

Func _check2($aArray, ByRef $iRow)
  If _check1($aArray) Then Return True
  ; Adapt $iRow :
  If $iRow < 1 Or $iRow > $aArray[0] Then $iRow = 1
  Return False
EndFunc   ;==>_check2

Func _check3($aArray, ByRef $iRow, ByRef $iColumn)
  If _check2($aArray, $iRow) Then Return True
  ; Check validity of GetA($aArray[$iRow], $iColumn) :
  If $iColumn < 1 Or $iColumn > $aArray[0] Then $iColumn = $aArray[0]
  Return False
EndFunc   ;==>_check3

Func _check4($aaArray, ByRef $iStart, ByRef $iEnd, ByRef $iColumn)
  If _check3($aaArray, $iStart, $iEnd) Then Return True
  ; Check $iColumn for all rows :
  If $iColumn < 0 Then $iColumn = 1
  If $iColumn = 0 Or @Compiled Then Return False
  Local $I
  For $I = $iStart To $iEnd
    If Not IsArray($aaArray[$I]) Then Return True
    If $iColumn > GetA($aaArray[$I], 0) Then Return True
  Next
  Return False
EndFunc   ;==>_check4

#EndRegion ; Safeguard boundaries and check array =================================

#Region ; Tests =================================================================

Func UseDroste()

  ; Build three dimensional Droste array:
  Local Enum $pN, $pNAME, $pBIRTHDAY, $pOCCUPATION, $pMARRIED, $pPLACES
  aD("John de Vries", "1986-12-23", "Retail manager", True, aRecite("Bremen|Zwolle"))
  aD("Christiane Delore", "1982-05-01", "Teacher", True, aRecite("New York|Miami|Emmen"))
  aD("Marijke van Dun", "unknown", "Carpenter", False, aRecite("Utrecht|London"))
  aD("Joost ten Velde", "1984-03-20", "Editor", False, aRecite("Amsterdam|Maastricht"))
  aD("Richard Nix", "1987-07-14", "Plumber", True, aRecite("Antwerpen|Paris"))
  Local $aPeople = aD()
  MsgBox(64, "Number of people", $aPeople[$pN])

  ; Search occupation of first teacher:
  Local $iFound = aSearch($aPeople, "Teacher", $pOCCUPATION)
  MsgBox(64, "Found", rRecite(aLeft($aPeople[$iFound], $pMARRIED), ", "))

  ; Retrieve teacher's places:
  MsgBox(64, "Teacher's places", rRecite(GetA($aPeople[$iFound], $pPLACES)))

  ; Update Marijke's birthday:
  SetA($aPeople[aSearch($aPeople, 'Marijke van Dun', $pNAME)], $pBIRTHDAY, "1983-04-25")

  ; Sort ascending on birthday:
  aCombSort($aPeople, $pBIRTHDAY, True)
  MsgBox(64, "From young to old", rRecite(aColumn($aPeople, $pNAME), ", "))

  ; More sort tests :
  Local $aR1 = aRound(aConcat(6, 5, 5, 5, 7))
  Local $nTest1 = aSimpleCombSort($aR1, True, True)
  MsgBox(64, "aSimpleCombSort. Descending numbers. nPass=" & $nTest1, rRecite($aR1))
  Local $nTest2 = aSimpleCombSort($aR1, False, True)
  MsgBox(64, "aSimpleCombSort. Ascending numbers. nPass=" & $nTest2, rRecite($aR1))
  Local $aR2 = aRecite("Portuguese|French|French|French|Spannish")
  Local $nTest3 = aSimpleCombSort($aR2, True, False)
  MsgBox(64, "aSimpleCombSort. Descending strings. nPass=" & $nTest3, rRecite($aR2))
  Local $nTest4 = aSimpleCombSort($aR2, False, False)
  MsgBox(64, "aSimpleCombSort. Ascending strings. nPass=" & $nTest4, rRecite($aR2))
  Local $aR0 = aNewArray(0)
  Local $nTest0 = aSimpleCombSort($aR0, True, True)
  MsgBox(64, "aSimpleCombSort. Zero elements. nPass=" & $nTest0, rRecite($aR0))

  ; StringFromArray and ArrayFromString:
  MsgBox(64, 'CompareArray', CompareArray($aPeople, ArrayFromString(StringFromArray($aPeople))))

  ; aStringFindAll :
  Local $S = "De kapper kapt knap maar de knecht van de knappe kapper kapt knapper dan de knappe kapper kappen kan"
  MsgBox(64, "Positions in string of 'kapper'", rRecite(aStringFindAll($S, "kapper"), ", "))

EndFunc   ;==>UseDroste

; UseDroste()

#EndRegion ; Tests =================================================================
