#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=aDrosteArray.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Droste array functions
#AutoIt3Wrapper_Res_Fileversion=2.0
#AutoIt3Wrapper_Res_LegalCopyright=Bert Kerkhof 2018-06-07 Apache 2.0 license
#AutoIt3Wrapper_Res_SaveSource=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 3 -w 4 -w 5
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 2 /reel
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv /rm
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; The latest version of this source is published at GitHub in
; repository "aDrosteArray".

#include-once
#include <AutoItConstants.au3> ; Delivered with AutoIT
#include <StringConstants.au3> ; Delivered with AutoIT
#include <FileConstants.au3> ; Delivered with AutoIT
#include <StringElaborate.au3> ; Published on GitHub, repository: "StringElaborate"
#include <FileJuggler.au3> ; Published on GitHub, repository: "FileJuggler"

; Tested with AutoIT v3.3.14.5 interpreter/compiler

; aDrosteArray module : basic array-in-array functions ========================
;    These differ from zero based array routines delivered with many
;    traditional computer interpreters and compilers. They also differ with the
;    AutoIt _array package:
;
; A. The place with index zero in aDrosteArray is reserved for a property that
;    reflects the total number of elements. The field is kept up-to-date by the
;    DrosteArray routines. So whether the user deletes an element with
;    aDelete() or inserts a element with aInsert(), the zero-element always
;    contains the total number of user-elements.
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
;    it afterwards in any matrix. This reduces the need to do double coding in
;    routines for both one- and two dimensional data input.
;
; D. With the aAdd and build functions, your program will never run stuck on
;    preprogrammed dimensional limitations or square sizes. Droste array's add
;    a level of abstraction. Focus on basics avoids a waste of subscript
;    numbers. Using lists in stead of single values will propel productivity.
;    Droste makes it work almost as pleasant as the use of strings.
;
;    The module is named after a dutch chocolate maker. For the recursive visual
;    effect, see:  https://en.wikipedia.org/wiki/Droste_effect
;
;       Bert Kerkhof
;
; Contents:
;    Create array's and add:
;       + aNewArray       + Creates a Droste array (one or two dimensions)
;       + aAdd            + Add value (single element or array) to an array
;       + aConcat         + Initialize an array, quickly assign values
;       + aD              + Add a row of multiple values to a 2dim array
;       + aRecite         + Place string values separated by '|' in array
;    Display array's:
;       + sRecite         + Place array values in a single line of text
;    Matrix operations on rows and columns:
;       + aLeft           + Copy the first items from an array
;       + aRight          + Copy the last items from an array
;       + aMid            + Copy items from an array
;       + aRound          + Converts an array of values to array of numbers
;       + aDec            + Convert string of hex values to an array of decimal
;       + sHex            + Converts array of decimal to a string of hex values
;       + aColumn         + Copy single column from two-dimension array.
;       + aCompare        + Compare contents of two multi-dim arrays
;    Operations on matrix elements:
;       + GetA            + Retrieve value from a 2dim array
;       + SetA            + Sets value in a 2dim array
;       + IncrA           + Increments value in a 2dim array
;       + MaxA            + Returns the highest value in an array
;       + MinA            + Returns the lowest value in an array
;       + LastIndex       + Returns highest index number (length) of array
;    Search and find:
;       + aSearch         + Search element or row in array
;       + aFindAll        + Collect multiple index positions in an array
;       + aStringFindAll  + Collect in an array multiple string occurrences
;    Array modifications:
;       + aDelete         + Delete element or row in an array
;       + aInsert         + Insert element or row in an array
;       + aFlat           + Collect multiple values and arrays in single array
;       + aFill           + Fill array with value
;    Sort:
;       + aCombSort       + Sorts a two-dimensional array
;       + aSimpleCombSort + Sorts a one-dimensional array
;    Array conversion:
;       + aDroste         + Convert zero-based array to type Droste
;       + msBrace         + Provides for use of aDroste array in For..In..Next
;       + aSquare         + Converts two dims of a Droste array to zero-based
;       + StringUtfArray  + Converts one-dim array of utf16 values to string
;       + ArrayUtfString  + Converts string to one-dim array of utf16 values
;       + StringFromArray + Converts (multi dim) array to a single string
;       + ArrayFromString + Restores a (multi dim) array from a single string
;    Array store and restore:
;       + WriteArray      + Write a (multi dim) array to file
;       + ReadArray       + Read a (multi dim) array from file

; Author ....: Bert Kerkhof ( kerkhof.bert@gmail.com )
; Tested with: AutoIt version 3.3.14.5 and win10 / win7

#Region ; Create and build ====================================================

Func aNewArray($nDim1 = 0, $nDim2 = 0) ; Create a new array
  ; Syntax.........: aNewArray(Dim1 [, Dim2])
  ; Parameters ....: - Dim1 (optional) last index (length) in the first
  ;                    dimension. If omitted, array will have 0 elements
  ;                  - Dim2 (optional) of the second dimension.
  ;                    If omitted, a one-dimensional array is created
  ; Return values .: Array with the specified dimension(s)
  ; See also ......: aAdd aConcat aRecite

  Local $aR1[1] = [0], $aR2[1] = [0]
  _NewDim($aR1, $nDim1 + 1)
  $aR1[0] = $nDim1
  If @NumParams = 2 Then
    _NewDim($aR2, $nDim2 + 1)
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

  If _check1($aArray) Then Return _errmess("aAdd", @error, 0, $aArray[0])
  $aArray[0] += 1
  _NewDim($aArray, $aArray[0])
  $aArray[$aArray[0]] = $Row
  Return $aArray[0]
EndFunc   ;==>aAdd

Func aConcat($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0, $K = 0, $M = 0, $N = 0, $P = 0, $Q = 0, $R = 0)
  ; Description ...: Concatenates up to 16 values: numbers / strings / arrays or a mixture
  ;                : If a value is an array, the resulting array is multi-dimensional
  ; Syntax.........: Array($v1 [,$v2 [,.. [, $v16 ]]])
  ; Parameters ....: $v1  - First element of the array
  ;                  $v2  - [optional] Second element of the array
  ;                  ...
  ;                  $v16 - [optional] Sixteenth element of the array
  ; Returns .......: Array with values, the zero element contains the number of values
  ; See also ......: aRecite aFlat aNewArray aAdd

  Local $aArray = [@NumParams, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J, $K, $M, $N, $P, $Q, $R]
  Return $aArray
EndFunc   ;==>aConcat

Func aD($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0, $K = 0, $M = 0, $N = 0, $P = 0, $Q = 0, $R = 0)
  ; Description ...: Builds two dimensional array. Each row contains up to 16 values:
  ;                  numbers / strings / arrays or mixture. When called with zero
  ;                  parameters, the build-up is returned and re-initialized.
  ; Returns .......: Array with a two dimensional array of values
  ; Author ........: Bert Kerkhof
  Local $nParam = @NumParams
  Local Static $aSave = aNewArray()
  If $nParam Then
    Local $aB = [$nParam, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J, $K, $M, $N, $P, $Q, $R]
    aAdd($aSave, aLeft($aB, $nParam))
  Else
    Local $aR = $aSave
    $aSave = aNewArray()
    Return $aR
  EndIf
EndFunc   ;==>aD

Func aRecite(Const $rString, $Sep = '|')
  ; Convert a recite of values to array
  ; Also see ......: sRecite aDec
  Return StringSplit($rString, $Sep) ; ® Split on vertical bar (not on comma)
EndFunc   ;==>aRecite

#EndRegion ; Create and build =================================================

#Region ; Display =============================================================

Func sRecite(Const $aArray, $sSeparator = '|', $sPrefix = '', $sPostfix = '')
  ; Readable array values on a single text line
  ; Syntax.........: sRecite(aArray, Separator, QuoteFlag, Fetch)
  ; Parameters ....: aArray      - Series of values to put readable on a line
  ;                  sSeparator  - [Optional] String that separates values,
  ;                                default : vertical bar
  ;                  sPrefix     - [Optional] String will prefix each element
  ;                  sPostFix    - [Optional] String will postfix each element
  ; Returns        : Readable string

  If _check1($aArray) Then Return _errmess("sRecite", @error, 0, '')
  Local $sResult = ''
  For $sFragment In msBrace($aArray)
    $sResult &= $sPrefix & $sFragment & $sPostfix & $sSeparator
  Next
  Return StringTrimRight($sResult, StringLen($sSeparator)) ; ® readable result
EndFunc   ;==>sRecite

#EndRegion ; Display ==========================================================

#Region ; Matrix operations on rows and columns ===============================

Func aColumn($aaArray, $iColumn, $iStart = 1, $iEnd = 0)
  ; Create single dim array from column in a multi-dim array
  If _check4($aaArray, $iStart, $iEnd, $iColumn) Then Return _errmess("aColumn", @error, @extended, aNewArray())
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
  If _check1($aArray) Then Return _errmess("aLeft", @error, 0, aNewArray())
  $aArray[0] = Max(0, Min($aArray[0], $N))
  Return $aArray
EndFunc   ;==>aLeft

Func aRight($aIn, $N = 1)
  ; Returns the last $N elements of an array:
  If _check1($aIn) Then Return _errmess("aRight", @error, 0, aNewArray())
  $N = Max(0, Min($N, $aIn[0]))
  Local $I, $aOut = aNewArray($N), $nOffset = $aIn[0] - $N
  For $I = 1 To $N
    $aOut[$I] = $aIn[$I + $nOffset]
  Next
  Return $aOut
EndFunc   ;==>aRight

Func aMid($aIn, $iPos, $N = 999999999999)
  ; Returns $N elements of an array, starting at $iPos:
  If _check2($aIn, $iPos) Then Return _errmess("aMid", @error, 0, aNewArray())
  $N = Max(0, Min($N, $aIn[0] - $iPos + 1))
  Local $I, $aOut = aNewArray($N)
  For $I = 1 To $N
    $aOut[$I] = $aIn[$I + $iPos - 1]
  Next
  Return $aOut
EndFunc   ;==>aMid

Func aRound($aArray, $nDecimal = 0)
  ; Converts an array of values to type number
  If _check1($aArray) Then Return _errmess("aRound", @error, 0, aNewArray())
  Local $I
  For $I = 1 To $aArray[0]
    $aArray[$I] = Round($aArray[$I], $nDecimal)
  Next
  Return $aArray
EndFunc   ;==>aRound

Func aDec(Const $sString, $sSep = '|')
  ; Convert hexadecimal values held in a string to array of numbers
  ; Related .......: sHex
  Local $aR = StringSplit($sString, $sSep) ; Split on vertical bar char
  For $I = 1 To $aR[0]
    $aR[$I] = Dec($aR[$I])
  Next
  Return $aR
EndFunc   ;==>aDec

Func sHex(Const $aArray, $sSep = '|')
  ; Convert an array of decimal values to string of hex numbers
  ; Related .......: aDec
  Local $sOut = ''
  For $I = 1 To $aArray[0]
    $sOut &= Hex($aArray[$I]) & $sSep
  Next
  Return StringTrimRight($sOut, 1)
EndFunc   ;==>sHex

Func LastIndex(Const $aArray)
  ; Returns the highest index number of array
  ; Syntax.........: LastIndex(aArray)
  ; Return values .: Index number
  If _check1($aArray) Then Return _errmess("LastIndex", @error, 0, 0)
  Return $aArray[0]
EndFunc   ;==>LastIndex

Func MaxA(Const $aArray, $iStart = 1, $iEnd = 0)
  ; Returns the highest value held in the numeric array
  ; Syntax.........: MaxA($aArray[, $iStart = 1[, $iEnd = 0]])
  ; Parameters ....: $aArray  - Array to search
  ;                  $iStart  - [optional] Index of array to start searching
  ;                  $iEnd    - [optional] Index of array to stop searching
  ; Return values .: The maximum value
  ; Also see ......: aMin, Max

  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("MaxA", @error, 0, 0)
  Local $I, $iMaxIndex = $iStart
  For $I = $iStart + 1 To $iEnd
    If $aArray[$iMaxIndex] < $aArray[$I] Then $iMaxIndex = $I
  Next
  Return $aArray[$iMaxIndex]
EndFunc   ;==>MaxA

Func MinA(Const $aArray, $iStart = 1, $iEnd = 0)
  ; Returns the lowest value held in the numeric array.
  ; Syntax.........: MinA($aArray[, $iStart = 1[, $iEnd = 0]])
  ; Parameters ....: $aArray  - Array to search
  ;                  $iStart  - [optional] Index of array to start searching
  ;                  $iEnd    - [optional] Index of array to stop searching
  ; Return values .: The minimum value
  ; Also see ......: aMax, Min

  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("MinA", @error, 0, 0)
  Local $I, $iMaxIndex = $iStart
  For $I = $iStart + 1 To $iEnd
    If $aArray[$iMaxIndex] > $aArray[$I] Then $iMaxIndex = $I
  Next
  Return $aArray[$iMaxIndex]
EndFunc   ;==>MinA

#EndRegion ; Matrix operations on rows and columns ============================

#Region ; Operations on matrix elements =======================================

Func GetA(Const $aR, $iP = 1)
  ; Get value from a (2dim) array
  ; Comment .......: Useful for Droste 2dim array
  ; Also see ......: aSet, aIncr
  ; Author ........: Thanks Wouter Bos
  If _check1($aR) Or $iP < 0 Or $iP > $aR[0] Then Return _errmess("GetA", @error, 0, '')
  Return $aR[$iP]
EndFunc   ;==>GetA

Func SetA(ByRef $aR, $iP = 1, $Val = 0)
  ; Set value in a (2dim) array
  ; Comment .......: Useful for Droste 2dim array
  ; Also see ......: aGet, aIncr
  If _check1($aR) Or $iP < 0 Or $iP > $aR[0] Then Return _errmess("SetA", @error, 0, '')
  $aR[$iP] = $Val
  Return $aR[$iP]
EndFunc   ;==>SetA

Func IncrA(ByRef $aR, $iP = 1, $N = 1)
  ; Increments value in a (2dim) array with +/- $N
  ; Comment .......: Useful for Droste 2dim array
  ; Also see ......: aGet, aSet
  If _check1($aR) Or $iP < 0 Or $iP > $aR[0] Then Return _errmess("IncrA", @error, 0, '')
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

#EndRegion ; Operations on matrix elements ====================================

#Region ; Search and find =====================================================

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

  If _check4($aArray, $iStart, $iEnd, $iColumn) Then Return _errmess("aSearch", @error, @extended, 0)
  Local $I, $nOffset = 0
  If $iColumn Then
    $aArray = aColumn($aArray, $iColumn, $iStart, $iEnd)
    If @error Then Return SetError(@error, @extended, 0)
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

  Local $I, $aR = aNewArray()
  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("aFindAll", @error, 0, $aR)
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
  Local $nOffset, $aR = aNewArray()
  While True
    $nOffset = StringInStr($sString, $sSearch, $STR_NOCASESENSE, 1, $iStart, $iEnd - $iStart + 1)
    If $nOffset = 0 Then ExitLoop
    aAdd($aR, $nOffset)
    $iStart = $nOffset + 1
  WEnd
  Return $aR
EndFunc   ;==>aStringFindAll

#EndRegion ; Search and find ==================================================

#Region ; Array modification ==================================================

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

  If _check2($aArray, $iPos) Then Return _errmess("aDelete", @error, 0, $aArray)
  $N = Min($N, $aArray[0] - $iPos + 1)
  If $iPos < 1 Or $N < 1 Then Return
  Local $J, $I = $aArray[0] - $N
  For $J = $iPos To $I
    $aArray[$J] = $aArray[$J + $N]
  Next
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

  If _check1($aArray) Then Return _errmess("aInsert", @error, 0, $aArray)
  If $iPos < 1 Then $iPos = 1
  If $iPos > $aArray[0] + 1 Then $iPos = $aArray[0] + 1

  Local $J, $vPos
  Local $nSize = IsArray($Value) ? $Value[0] : 1
  $aArray[0] += $nSize ; Increase length of the array
  _NewDim($aArray, $aArray[0])
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

Func aFlat($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0, $K = 0, $M = 0, $N = 0, $P = 0, $Q = 0, $R = 0)
  ; Description ...: Serially concatenates up to 16 array's / numbers / strings
  ; Syntax.........: aSerialConcat(aArray1, aArray1, Value .. $aArrayN)
  ; Returns .......: One-dimensional array of values
  ; See also ......: aRecite aConcat aNewArray aAdd
  ; Author ........: Bert Kerkhof

  Local $L, $aReturn = aNewArray()
  Local $aArrayB = [@NumParams, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J, $K, $M, $N, $P, $Q, $R]
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
  If _check1($aArray) Then Return _errmess("aFill", @error, 0, 0)
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

#EndRegion ; Array modification ==================================================

#Region ; Array conversion back and forth =====================================

Func aDroste(Const $aArray) ; Convert 1dim or 2dim array to type Droste
  ; Syntax.........: aDroste(aArray)
  ; Parameter......: aArray - Array to be converted
  ; Comment........: A first element is added with the length of the array
  ; Return value...: The aDrosted array
  ; Remarks........: Execution time 135 milliseconds on quad computer for
  ;                  100.000 elements array
  If Not IsArray($aArray) Then Return _errmess("aDroste", 1, 0, aNewArray())
  Local $I, $nRow = UBound($aArray, 1), $aRet = aNewArray($nRow)
  $aRet[0] = $nRow
  Switch UBound($aArray, $UBOUND_DIMENSIONS)
    Case 1
      For $I = 1 To $nRow
        $aRet[$I] = $aArray[$I - 1]
      Next
    Case 2
      Local $J, $nCol = UBound($aArray, 2), $aRow = aNewArray($nCol)
      For $I = 1 To $nRow
        $aRow[0] = $nCol
        For $J = 1 To $nCol
          $aRow[$J] = $aArray[$I - 1][$J - 1]
        Next
        $aRet[$I] = $aRow
      Next
    Case Else
      Return _errmess("aDroste", 4, 0, aNewArray())
  EndSwitch
  Return $aRet
EndFunc   ;==>aDroste

Func msBrace(Const $Array)
  ; With the use of msBrace() the aDroste array is able to function as a
  ; collection in the microsoft syntax: 'For $Var In $Array'
  ;
  ; Parameter......: $Array - The aDroste array to convert.
  ; Returns........: The zero based array. If the array is empty, a square
  ;                  two-dimensional array is returned that will exit the
  ;                  loop and continues as usual. See: help file =>
  ;                  Language Reference => Loop Statements => For..In..Next
  ; Comment........: The number of elements in the constructed zero
  ;                  base array can be obtained with Ubound($aArray, 1)

  Local $dummy[1][1] ; To signal an exit from loop
  If Not IsArray($Array) Then Return _errmess("msBrace", @error, 0, $dummy)
  If $Array[0] = 0 Then Return $dummy
  Local $zR[$Array[0]] ; 1dim array
  For $I = 1 To $Array[0] ; Transfer values:
    $zR[$I - 1] = $Array[$I]
  Next
  Return $zR
EndFunc   ;==>msBrace

Func zSquare(Const $aArray)
  ; Convert a one- or two-dimensional aDroste to zero based square array
  ; Syntax.........: zSquare(aArray)
  ; Parameter......: $aArray - The aDroste array to be converted.
  ; Return value...: The zero based square array.
  ; Comment1.......: With the use of zSquare() the aDroste array is able to excel
  ;                  between zero based array syntax.
  ; Comment2.......: The number of rows and columns in the constructed square array
  ;                  can be obtained with Ubound($aArray, 1) and Ubound($aArray, 2).

  Local $I, $Array, $nColumn = 0
  If Not IsArray($aArray) Then Return _errmess("zSquare", @error, 0, aNewArray(0))
  For $I = 1 To $aArray[0] ; Obtain# of columns:
    $Array = $aArray[$I]
    If IsArray($Array) Then $nColumn = Max($nColumn, $Array[0])
  Next
  If $nColumn = 0 Then Return msBrace($aArray)
  Local $J, $R[$aArray[0]][$nColumn] ; 2dim array
  For $I = 1 To $aArray[0] ; Transfer values:
    $Array = $aArray[$I]
    For $J = 1 To $Array[0]
      $R[$I - 1][$J - 1] = $Array[$J]
    Next
  Next
  Return $R
EndFunc   ;==>zSquare

Func StringUtfArray($aArray, $iStart = 1, $iEnd = 0)
  ; Convert an array of utf16 numbers into a string
  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("StringFromArray", @error, 0, '')
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

#EndRegion ; Array conversion back and forth ==================================

#Region ; Sort ================================================================

Func NumberCompare($N1, $N2)
  ; helper function for aCombSort :
  If $N1 > $N2 Then Return 1
  If $N1 < $N2 Then Return -1
  Return 0
EndFunc   ;==>NumberCompare

Func Lif($Logic, $P1, $P2 = '') ; Select value depending on condition
  ; helper function.
  ; Warning: Although logic only chooses one, the function will activate
  ; both assignment choices. Use the 'ternari' syntax to avoid this.
  Return $Logic ? $P1 : $P2
EndFunc   ;==>Lif

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

  If _check4($aArray, $iStart, $iEnd, $iColumn) Then Return _errmess("aCombSort", @error, @extended, 0)
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

  If _check3($aArray, $iStart, $iEnd) Then Return _errmess("aSimpleCombSort", @error, 0, 0)
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

#EndRegion ; Sort =============================================================

#Region ; Array <=> String and file  ==========================================

Func StringFromArray($aArray, $S = ChrW(9))
  ; Restore (multi dim) array from a single string
  ; Parameters ..: $aArray : (multi dimensional aDroste) array
  ;                $S      : Unique separator char
  ;                          Choose carefully in case of binary data
  ;                          Default: TAB
  ; Returns .....: $String : Result
  ; Author ......: Bert Kerkhof

  If _check1($aArray) Then Return _errmess("StringFromArray", @error, 0, False)
  If StringLen($S) = 1 Then $S &= '0'
  Local $I, $String = ''
  For $aR In msBrace($aArray)
    If IsArray($aR) Then
      Local $SN = StringLeft($S, 1) & ChrW(AscW(StringRight($S, 1)) + 1)
      $String &= ($aR[0] > 1) ? $S : $S & $SN
      $String &= StringFromArray($aR, $SN) ; Recurse
    Else
      If $I > 1 Then $String &= $S
      $String &= $aR
    EndIf
  Next
  Return $String
EndFunc   ;==>StringFromArray

Func ArrayFromString($String, $S = '')
  ; Store the contents of (multi dim) array in a single string
  ; Parameters ..: $String : Previous encoded data, see StringFromArray()
  ;                $S      : Two unique separator chars read from string,
  ;                          blank at start-up
  ; Returns .....: $aR     : Result array
  ; Author ......: Bert Kerkhof

  If StringLen($S) < 2 Then $S = StringLeft($String, 2)
  If $S = StringLeft($String, 2) Then $String = StringMid($String, 3)
  Local $aSp = StringSplit($String, $S, $STR_ENTIRESPLIT)
  Local $SN = StringLeft($S, 1) & ChrW(AscW(StringRight($S, 1)) + 1)
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

  If _check1($aArray1) Or _check1($aArray2) Then Return _errmess("aCompare", @error, 0, False)
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

#EndRegion ; Array <=> String and file  =======================================

#Region ; Maintainance and check array boundaries =============================

Func _NewDim(ByRef $aArray, $nDim)
  If $nDim < UBound($aArray) Then Return
  ; 16, 32, 64, 128, 256, 512, 1024, 2048 ..
  ReDim $aArray[2 ^ Ceiling(Log(Max($nDim, 16)) / Log(2)) + 1]
EndFunc   ;==>_NewDim

Func _errmess($sOrigin, $error, $extended, $Rvalue)
  If @Compiled Then Return SetError($error, $extended, $Rvalue)
  Local $aS1 = aRecite('Not an array|Not an array in second dim|An array in second dim is too short|More than two dimensions')
  Local $S1 = ($error > 0 And $error < 4) ? @CRLF & $aS1[$error] : ""
  Local $S2 = $extended ? @CRLF & "Element number " & $extended : ""
  MsgBox(64, FileBase(FileName(@ScriptName)), 'Error in function ' & $sOrigin & $S1 & $S2)
  Exit
  Return SetError($error, $extended, $Rvalue)
EndFunc   ;==>_errmess

Func _check1($aArray)
  ; Check validity of $aArray :
  If Not IsArray($aArray) Then Return SetError(1)
  Return False
EndFunc   ;==>_check1

Func _check2($aArray, ByRef $iRow)
  If _check1($aArray) Then Return SetError(@error)
  ; Adapt $iRow :
  If $iRow < 1 Or $iRow > $aArray[0] Then $iRow = 1
  Return False
EndFunc   ;==>_check2

Func _check3($aArray, ByRef $iRow, ByRef $iColumn)
  If _check2($aArray, $iRow) Then Return SetError(@error)
  ; Check validity of ($aArray[$iRow])[$iColumn] :
  If $iColumn < 1 Or $iColumn > $aArray[0] Then $iColumn = $aArray[0]
  Return False
EndFunc   ;==>_check3

Func _check4($aaArray, ByRef $iStart, ByRef $iEnd, ByRef $iColumn)
  If _check3($aaArray, $iStart, $iEnd) Then Return SetError(@error)
  ; Check $iColumn for all rows :
  If $iColumn < 0 Then $iColumn = 1
  If $iColumn = 0 Or @Compiled Then Return False
  Local $I
  For $I = $iStart To $iEnd
    If Not IsArray($aaArray[$I]) Then Return SetError(2, $I)
    If $iColumn > ($aaArray[$I])[0] Then Return SetError(3, $I)
  Next
  Return False
EndFunc   ;==>_check4

#EndRegion ; Maintainance and check array boundaries ==========================

#Region ; Tests ===============================================================

Func UseDroste()

  ; Build three dimensional Droste array:
  Local Enum $pNAME = 1, $pBIRTHDAY, $pOCCUPATION, $pMARRIED, $pPLACES
  aD("John de Vries", "1986-12-23", "Retail manager", True, aRecite("Bremen|Zwolle"))
  aD("Christiane Delore", "1982-05-01", "Teacher", True, aRecite("New York|Miami|Emmen"))
  aD("Marijke van Dun", "unknown", "Carpenter", False, aRecite("Utrecht|London"))
  aD("Joost ten Velde", "1984-03-20", "Editor", False, aRecite("Amsterdam|Maastricht"))
  aD("Richard Nix", "1987-07-14", "Plumber", True, aRecite("Antwerpen|Paris"))
  Local $aPeople = aD()
  MsgBox(64, "Number of people", $aPeople[0])

  ; Search occupation of first teacher:
  Local $iFound = aSearch($aPeople, "Teacher", $pOCCUPATION)
  MsgBox(64, "Found", sRecite(aLeft($aPeople[$iFound], $pMARRIED), ", "))

  ; Retrieve teacher's places:
  MsgBox(64, "Teacher's places", sRecite($aPeople[$iFound])[$pPLACES])

  ; Update Marijke's birthday:
  SetA($aPeople[aSearch($aPeople, 'Marijke van Dun', $pNAME)], $pBIRTHDAY, "1983-04-25")

  ; Sort ascending on birthday:
  aCombSort($aPeople, $pBIRTHDAY, True)
  MsgBox(64, "From young to old", sRecite(aColumn($aPeople, $pNAME), ", "))

  ; More sort tests :
  Local $aR1 = aRound(aConcat(6, 5, 5, 5, 7))
  Local $nTest1 = aSimpleCombSort($aR1, True, True)
  MsgBox(64, "aSimpleCombSort. Descending numbers. nPass=" & $nTest1, sRecite($aR1))
  Local $nTest2 = aSimpleCombSort($aR1, False, True)
  MsgBox(64, "aSimpleCombSort. Ascending numbers. nPass=" & $nTest2, sRecite($aR1))
  Local $aR2 = aRecite("Portuguese|French|French|French|Spannish")
  Local $nTest3 = aSimpleCombSort($aR2, True, False)
  MsgBox(64, "aSimpleCombSort. Descending strings. nPass=" & $nTest3, sRecite($aR2))
  Local $nTest4 = aSimpleCombSort($aR2, False, False)
  MsgBox(64, "aSimpleCombSort. Ascending strings. nPass=" & $nTest4, sRecite($aR2))
  Local $aR0 = aNewArray()
  Local $nTest0 = aSimpleCombSort($aR0, True, True)
  MsgBox(64, "aSimpleCombSort. Zero elements. nPass=" & $nTest0, sRecite($aR0))

  ; StringFromArray and ArrayFromString:
  MsgBox(64, 'CompareArray', CompareArray($aPeople, ArrayFromString(StringFromArray($aPeople))))

  ; aStringFindAll :
  Local $S = "De kapper kapt knap maar de knecht van de knappe kapper kapt knapper dan de knappe kapper kappen kan"
  MsgBox(64, "Positions in string of 'kapper'", sRecite(aStringFindAll($S, "kapper"), ", "))

EndFunc   ;==>UseDroste

; UseDroste()

Func TestNewDim()
  Local $aR = aNewArray()
  cc('aNewArray, inital dimension: ' & UBound($aR))
  aInsert($aR, 1, aNewArray(16))
  cc('After adding 16 values: ' & UBound($aR))
  aAdd($aR, 1)
  cc('After adding the 17th value: ' & UBound($aR))
EndFunc   ;==>TestNewDim
; TestNewDim()

Func TestzSquareAndaDroste()
  Local $aaT = aConcat(aRecite("one|two|three"), aRecite("four|five|six"))
  Local $R2 = zSquare($aaT) ; Create zero based array
  cc('zRow1 0: ' & $R2[0][0] & " " & $R2[0][1] & " " & $R2[0][2])
  cc('zRow2 1: ' & $R2[1][0] & " " & $R2[1][1] & " " & $R2[1][2])
  Local $aaR = aDroste($R2) ; And convert back to array-in-array
  cc('1th element: ' & sRecite($aaR[1]))
  cc('2th element: ' & sRecite($aaR[2]))

  $aaT = aConcat("seven", "eight", "nine")
  $R2 = zSquare($aaT)
  cc('zRow3: ' & $R2[0] & " " & $R2[1] & " " & $R2[2])
  cc("Result3: " & sRecite(aDroste($R2)))
EndFunc   ;==>TestzSquareAndaDroste
; TestzSquareAndaDroste()

#EndRegion ; Tests ============================================================
