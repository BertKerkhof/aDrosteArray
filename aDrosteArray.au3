#Region ;**** Directives created by AutoIt3Wrapper_GUI ***
#AutoIt3Wrapper_Outfile=aDrosteArray.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Droste array functions
#AutoIt3Wrapper_Res_Fileversion=2.0
#AutoIt3Wrapper_Res_LegalCopyright=Bert Kerkhof 2019-11-07 Apache 2.0 license
#AutoIt3Wrapper_Res_SaveSource=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 3 -w 4 -w 5
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 2 /reel
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv /rm
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ***

; Version 4
; The latest version of this source is published at GitHub in the
; BertKerkhof repository "aDrosteArray".

#include-once
#include <AutoItConstants.au3>; Delivered with AutoIT
#include <StringConstants.au3>; Delivered With AutoIT
#include <FileConstants.au3>; Delivered With AutoIT
#include <StringSupport.au3>; GitHub published by Bert Kerkhof

; Author: Bert Kerkhof ( kerkhof.bert@gmail.com )
; Tested with AutoIT v3.3.14.5 interpreter/compiler and win10


; aDrosteArray module : basic array-in-array functions ================
;    These differ from zero based array routines delivered with many
;    traditional computer interpreters and compilers. They also differ with
;    the AutoIt _array package:
;
; A. The place with index zero in aDrosteArray is reserved for a property
;    that reflects the total number of elements. The field is kept
;    up-to-date by DrosteArray routines. So whether the user deletes an
;    element with aDelete() or inserts an element with aInsert(), the
;    zero-element always contains the total number of user-elements.
;
;    With a DrosteArray, the user does not have to calculate the number of
;    user elements with +1 nor the end-argument in a for-next loop with -1
;    anymore. This reduces the number of runtime out-of-bound messages. The
;    the total number of user-elements for every array is close at hand in
;    the aArray[0] element. With less worry about these details you can
;    concentrate on design and efficiency of the program.
;
; B. Droste multi-dimensional arrays are more versatile than square 2D
;    arrays. The option of different row sizes is more adapt to many
;    logistics which software designers face. Adding elements and specific
;    dimensions reflect real life situations which asks for detail.
;
; C. Many available routines operate on a single dimension array only. The
;    array-in-array design makes it easy to retrieve a single array and
;    restore it afterwards in any matrix. This reduces the need to do double
;    coding in routines for both one- and two dimensional data input.
;
; D. With the aAdd and build functions, your program will never run stuck on
;    preprogrammed dimensional limitations or square sizes. Droste array's
;    add a level of abstraction. Focus on basics avoids a waste of subscript
;    numbers. Using lists instead of single values will propel productivity.
;    Droste makes it work almost as pleasant as the use of strings.
;
;    The module is named after a Dutch chocolate maker. For the recursive
;    visual effect, see:  https://en.wikipedia.org/wiki/Droste_effect
;
;       Bert Kerkhof
;
; Contents:
;    Create array's and build:
;       + aNew            + Creates a Droste array (one or two dimensions)
;       + aAdd            + Add value (single element or array) to an array
;       + Array           + Initialize an array, quickly assign values
;       + aD              + Add a row of multiple values to a 2dim array
;       + aRecite         + Place string values separated by '|' in array
;    Display array's:
;       + sRecite         + Place array values in a single line of text
;    Operations on matrix elements:
;       + SetA            + Sets value in a 2dim array
;       + IncrA           + Increments value in a 2dim array
;       + MaxA            + Returns the highest value in an array
;       + MinA            + Returns the lowest value in an array
;       + LastIndex       + Returns highest index number (length) of array
;    Operations on rows and matrix:
;       + aLeft           + Copy the first items from an array
;       + aRight          + Copy the last items from an array
;       + aMid            + Copy items from an array
;       + aColumn         + Copy single column from two-dimension array.
;    Search and find:
;       + aSearch         + Search element or row in array
;       + aFindAll        + Collect multiple index positions in an array
;       + aStringLocate   + Collect in an array multiple string occurrences
;    Array modifications:
;       + aDelete         + Delete element or row in an array
;       + aInsert         + Insert element or row in an array
;       + aFlat           + Collect values and arrays in single array
;       + aFill           + Fill array with value
;    Sort:
;       + aCombSort       + Sorts a two-dimensional array
;       + aSimpleCombSort + Sorts a one-dimensional array
;    Array type conversion:
;       + aDroste         + Convert zero-based array to type Droste

; Version3 to version4 changeLog:
; The aim of version4 is a return to basics for educational purpose.
;   Less used functions aRound(), aDec(), aHex(), aCompare(),
;     msBrace(), aSquare(), StringUtfArray() ArrayUtfString(),
;     StringFromArray(), ArrayFromString(), WriteArray() and
;     ReadArray() are discontinued.
;   The GetA() function is surpassed by the ($Array[$I])[$J] construct.
;   aNewArray() is renamed to aNew().
;     aNew() has one parameter, supports single-dimensioned arrays only.
;   aConcat() is renamed to Array().
;   sRecite() now can have a numeric value as separator parameter.
;     In that case the outputted array elements will be right-justified.

; Version2 to version3 changeLog
;   aStringFindAll() is renamed to aStringLocate().


; Create and build ====================================================

; #FUNCTION#
; Name ..........: aNew
; Description ...: Create a new Droste array with the specified dimension.
; Syntax ........: aNew([$nDim = 0])
; Parameters ....: $nDim ... [optional] number of elements the array
;                            contains. By default the array has zero
;                            elements.
; Returns .......: Array with the specified dimension
; Author ........: Bert Kerkhof

Func aNew(Const $nDim = 0) ; Create a new array
  ; Syntax.........: aNew([$nDim = 0])
  ; Parameters ....: $nDim .. [optional] last index (length) in the first
  ;                           dimension. If omitted, array will have
  ;                           zero elements.
  ; Returns .......: Array with the specified dimension
  ; See also ......: aAdd aConcat aRecite

  Static Local $A[16] ; Optimized for execution speed
  If $nDim Then
    Local $B[$nDim + 1]
    $B[0] = $nDim
    Return $B
  EndIf
  Return $A
EndFunc   ;==>aNew

; #FUNCTION#
; Name ..........: aAdd
; Description ...: Add a value at the end of an array.
; Syntax ........: aAdd(Byref $aArray, Const $Value)
; Parameters ....: $aArray ... [in/out] array to extend.
;                  $Value .... [const] value to add.
; Returns .......: The new length of the array
;                  The extended array is passed by reference
; Comment .......: Because the memory allocation is fast, Add is an
;                  efficient way to build your custom sized array's from
;                  scratch
; Author ........: Bert Kerkhof

Func aAdd(ByRef $aArray, Const $Row)
  If _bounds1($aArray) Then Return _errmess("aAdd", @error, 0, 0)
  $aArray[0] += 1
  _NewDim($aArray)
  $aArray[$aArray[0]] = $Row
  Return $aArray[0]
EndFunc   ;==>aAdd

; #FUNCTION#
; Name ..........: Array
; Description ...: Concatenates up to 16 values.
; Syntax ........: Array([$A = 0[, $B = 0[, $C = 0[, $D = 0[, $E = 0[, $F = 0[, $G = 0[, $H = 0[, $I = 0[, $J = 0[,
;                  $K = 0[, $M = 0[, $N = 0[, $P = 0[, $Q = 0[, $R = 0]]]]]]]]]]]]]]]])
; Parameters ....: $A to $R....[optional] up to 16 values:
;                              numbers, strings..
; Returns .......: Array with values, the zero element contains the number
;                  of values
; Author ........: Bert Kerkhof

Func Array($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0, $K = 0, $M = 0, $N = 0, $P = 0, $Q = 0, $R = 0)
  Local $aArray = [@NumParams, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J, $K, $M, $N, $P, $Q, $R]
  Return $aArray
EndFunc   ;==>Array

; #FUNCTION#
; Name ..........: aD
; Description ...: Build a two-dimensional array
; Syntax ........: aD([$A = 0[, $B = 0[, $C = 0[, $D = 0[, $E = 0[, $F = 0[, $G = 0[, $H = 0[, $I = 0[, $J = 0[, $K = 0[,
;                  $M = 0[, $N = 0[, $P = 0[, $Q = 0[, $R = 0]]]]]]]]]]]]]]]])
; Parameters ....: $A to $R...... [optional] up to 16 values:
;                                 numbers / strings..
;                  When called with zero parameters, the build-up is
;                  returned and re-initialized.
; Returns .......: Two dimensional array
; Comment .......: This is another efficient way to quickly embed data in
;                  your source code
; Author ........: Bert Kerkhof

Func aD($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0, $K = 0, $M = 0, $N = 0, $P = 0, $Q = 0, $R = 0)
  Local $nParam = @NumParams
  Local Static $aSave = aNew()
  Switch $nParam
    Case 0
      Local $aResult= $aSave
      $aSave = aNew(0)
      Return $aResult
    Case Else
      Local $aB = [$nParam, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J, $K, $M, $N, $P, $Q, $R]
      aAdd($aSave, aLeft($aB, $nParam))
  EndSwitch
EndFunc   ;==>aD

; #FUNCTION#
; Name ..........: aRecite
; Description ...: Convert a recite of values to array
; Syntax ........: aRecite(Const $rString[, $Sep = '|'])
; Parameters ....: $rString .. [const] recited values in a string.
;                  $Sep ...... [optional] character to separate the values.
;                              Default is '|'.
; Returns .......: Array of values
; Author ........: Bert Kerkhof

Func aRecite(Const $rString, Const $Sep = '|')
  Return StringSplit($rString, $Sep) ; ® Split on vertical bar (not on comma)
EndFunc   ;==>aRecite


; Display =============================================================

; #FUNCTION#
; Name ..........: sRecite
; Description ...: Present array values in a single string
; Syntax ........: sRecite(Const $aArray[, $Sep = '|'[, $sPrefix = ''[, $sPostfix = '']]])
; Parameters ....: $aArray ... Series of values to display readable on a line
;                  $Sep ...... [optional] If $Sep is a string, multiple
;                              $aArray elements (if any) will be separated
;                              by this string.
;                              If $Sep is a number, the elements are right-
;                              justified with $Sep column width.
;                              Default is "|".
;                  $sPrefix ..... [optional] string will prefix each element
;                  $sPostFix .... [optional] string will postfix each element
; Returns        : Readable string
; Author ........: Bert Kerkhof

Func sRecite(Const $aArray, Const $Sep = '|', $sPrefix = "", $sPostfix = "")
  If _bounds1($aArray) Then Return _errmess("sRecite", @error, 0, '')
  Local $sOut = ""
  If IsNumber($Sep) Then
    For $I = 1 to $aArray[0]
      $sOut &= sJustR($sPreFix & $aArray[$I] & $sPostfix, $Sep)
    Next
    Return $sOut ; ® readable result
  EndIf
  For $I = 1 to $aArray[0]
    $sOut &= $sPrefix & $aArray[$I] & $sPostfix & $Sep
  Next
  Return StringTrimRight($sOut, StringLen($Sep))
EndFunc   ;==>sRecite


; Operations on elements ==============================================

; #FUNCTION#
; Name ..........: SetA
; Description ...: Set value in a two-dimensional array
; Syntax ........: SetA(Byref $aArray[, $iP = 1[, $Val = 0]])
; Parameters ....: $aArray ... [in/out] row of a two dimensional
;                              array-in-array
;                  $iP ....... [optional] index to the second dimension
;                              By default element number 1 in the second
;                              dimension
;                  $Val ...... [optional] value to set. Default is 0.
; Return value ..: Updated single element.
;                  The updated row is return by reference
; Author ........: Bert Kerkhof

Func SetA(ByRef $aArray, Const $iP = 1, Const $Val = 0)
  If _bounds1($aArray) Or $iP < 0 Or $iP > $aArray[0] Then Return _errmess("SetA", @error, 0, '')
  $aArray[$iP] = $Val
  Return $aArray[$iP]
EndFunc   ;==>SetA

; #FUNCTION#
; Name ..........: IncrA
; Description ...: Increment value in a two-dimensional array
; Syntax ........: IncrA(Byref $aArray[, $iP = 1[, $N = 1]])
; Parameters ....: $aArray ... [in/out] row of a two dimensional
;                              array-in-array
;                  $iP ....... [optional] index to the second dimension
;                              By default element number 1 in the second
;                              dimension
;                  $N ........ [optional] increment. Default is 1.
; Return value ..: Updated single element.
;                  The updated row is return by reference
; Author ........: Bert Kerkhof

Func IncrA(ByRef $aArray, Const $iP = 1, Const $N = 1)
  ; Increments value in a (2dim) array with +/- $N
  ; Comment .......: Useful for Droste 2dim array
  ; Also see ......: aGet, aSet

  If _bounds1($aArray) Or $iP < 0 Or $iP > $aArray[0] Then Return _errmess("IncrA", @error, 0, '')
  $aArray[$iP] += $N
  Return $aArray[$iP]
EndFunc   ;==>IncrA

; #FUNCTION#
; Name ..........: Max
; Description ...: Compare numbers and return the highest one.
; Syntax ........: Max(Const $N1, Const $N2)
; Parameters ....: $N1, $N2 ... [const] numbers to compare
; Returns .......: Highest value.
; Author ........: Bert Kerkhof

Func Max(Const $N1, Const $N2)
  Return $N1>$N2 ? $N1 : $N2
EndFunc   ;==>Max

; #FUNCTION#
; Name ..........: Min
; Description ...: Compare numbers and return the lowest one
; Syntax ........: Min(Const $N1, Const $N2)
; Parameters ....: $N1, $N2 ... [const] numbers to compare
; Returns .......: Lowest value
; Author ........: Bert Kerkhof

Func Min(Const $N1, Const $N2)
  Return $N1>$N2 ? $N2 : $N1
EndFunc   ;==>Min

; #FUNCTION#
; Name ..........: MaxA
; Description ...: Return the highest value of a numeric array.
; Syntax ........: MaxA(Const $aArray[, $iStart = 1[, $iEnd = 0]])
; Parameters ....: $aArray ........ Array to evaluate
;                  $iStart, $End .. [optional] range of array elements
; Returns .......: Maximum value
; Author ........: Bert Kerkhof

Func MaxA(Const $aArray, Const $iStart = 1, $iEnd = 0)
  If $iEnd = 0 Then $iEnd = $aArray[0]
  If _bounds2($aArray, $iStart, $iEnd) Then Return _errmess("MaxA", @error, 0, 0)
  Local $iMaxIndex = $iStart
  For $I = $iStart + 1 To $iEnd
    If $aArray[$iMaxIndex] < $aArray[$I] Then $iMaxIndex = $I
  Next
  Return $aArray[$iMaxIndex]
EndFunc   ;==>MaxA

; #FUNCTION#
; Name ..........: MinA
; Description ...: Return the lowest value of a numeric array.
; Syntax ........: MinA(Const $aArray[, $iStart = 1[, $iEnd = 0]])
; Parameters ....: $aArray ........ Array to evaluate
;                  $iStart, $End .. [optional] range of array elements
; Returns .......: Minimum value
; Author ........: Bert Kerkhof

Func MinA(Const $aArray, Const $iStart = 1, $iEnd = 0)
  If $iEnd = 0 Then $iEnd = $aArray[0]
  If _bounds2($aArray, $iStart, $iEnd) Then Return _errmess("MinA", @error, 0, 0)
  Local $iMaxIndex = $iStart
  For $I = $iStart + 1 To $iEnd
    If $aArray[$iMaxIndex] > $aArray[$I] Then $iMaxIndex = $I
  Next
  Return $aArray[$iMaxIndex]
EndFunc   ;==>MinA

; #FUNCTION#
; Name ..........: LastIndex
; Description ...: Return the highest index number of array.
; Syntax ........: LastIndex(Const $aArray)
; Parameters ....: $aArray..... [const] array to inspect.
; Returns .......: Index number
; Author ........: Bert Kerkhof

Func LastIndex(Const $aArray)
  If _bounds1($aArray) Then Return _errmess("LastIndex", @error, 0, 0)
  Return $aArray[0]
EndFunc   ;==>LastIndex


; Operations on rows and matrix =======================================

; #FUNCTION#
; Name ..........: aColumn
; Description ...: Create single dim array from column in a multi-dim array.
; Syntax ........: aColumn(Const $aaArray[, $iColumn = 1[, $iStart = 1[, $iEnd = 0]]])
; Parameters ....: $aaArray......... [const] two-dimensional array.
;                  $iColumn......... [optional] integer value. Default is 1.
;                  $iStart, $iEnd .. [optional] range of $aaArray rows.
; Returns .......: Array of column values
; Author ........: Bert Kerkhof

Func aColumn(Const $aaArray, Const $iColumn = 1, Const $iStart = 1, $iEnd = 0)
  If $iEnd = 0 Then $iEnd = $aaArray[0]
  If _bounds3($aaArray, $iColumn, $iStart, $iEnd) Then Return _errmess("aColumn", @error, @extended, aNew())
  Local $N = $iEnd - $iStart + 1, $nOffset = $iStart - 1
  Local $aArray, $aResult= aNew($N)
  For $I = 1 To $N
    $aArray = $aaArray[$I + $nOffset]
    $aResult[$I] = $aArray[$iColumn]
  Next
  Return $aResult
EndFunc   ;==>aColumn

; #FUNCTION#
; Name ..........: aLeft
; Description ...: Return the first $N elements of an array.
; Syntax ........: aLeft($aArray[, $N = 1])
; Parameters ....: $aArray ..... Source array.
;                  $N .......... [optional] number of elements to return.
;                                Default is 1.
; Returns .......: Array of selected values
; Author ........: Bert Kerkhof

Func aLeft($aArray, Const $N = 1)
  ; Returns the first $N elements of an array:
  If _bounds1($aArray) Then Return _errmess("aLeft", @error, 0, aNew())
  $aArray[0] = Max(Min($aArray[0], $N), 0)
  Return $aArray
EndFunc   ;==>aLeft

; #FUNCTION#
; Name ..........: aRight
; Description ...: Return the last $N elements of an array.
; Syntax ........: aRight(Const $aIn[, $N = 1])
; Parameters ....: $aIn ..... [const] source array.
;                  $N ....... [optional] number of elements to return.
;                             Default is 1: only the leftmost element.
; Returns .......: Array of selected values
; Author ........: Bert Kerkhof

Func aRight(Const $aIn, $N = 1)
  If _bounds1($aIn) Then Return _errmess("aRight", @error, 0, aNew())
  $N = Max(Min($aIn[0], $N), 0)
  Local $aOut = aNew($N), $nOffset = $aIn[0] - $N
  For $I = 1 To $N
    $aOut[$I] = $aIn[$I + $nOffset]
  Next
  Return $aOut
EndFunc   ;==>aRight

; #FUNCTION#
; Name ..........: aMid
; Description ...: Return selection of array elements.
; Syntax ........: aMid(Const $aIn, $I[, $N = 2147483647])
; Parameters ....: $aIn .... [const] source array
;                  $I ...... First element to select
;                  $N ...... [optional] number of elements to select.
;                            Default is the range from $I upto last element.
; Returns .......: Array of selected values
; Author ........: Bert Kerkhof

Func aMid(Const $aIn, $I, $N = 2147483647) ; Max 32 bit integer (2**31 - 1)
  ; Returns $N elements of an array, starting at $I:
  If _bounds1($aIn) Then Return _errmess("aMid", @error, 0, aNew())
  $I = Max(1, $I) ; Prevent zero and negative index
  $N = Max(0, Min($N, $aIn[0] - $I + 1)) ; Max 0 for $N = 0
  Local $aOut = aNew($N)
  For $J = 1 To $N
    $aOut[$J] = $aIn[$I + $J - 1]
  Next
  Return $aOut
EndFunc   ;==>aMid


; Search and find =====================================================

; #FUNCTION#
; Name ..........: aSearch
; Description ...: Search an element in one- or two-dimensional array.
; Syntax ........: aSearch(Const $aArray, Const $Value[, $iColumn = 0[, $iStart = 1[, $iEnd = 0]]])
; Parameters ....: $aArray ......... Array to search in
;                  $Value .......... Element to search
;                  $iColumn ........ Zero for 1dim array
;                                    2dim: column to search
;                  $iStart, $iEnd .. [optional] range of array elements
; Returns .......: Found : Index number of the string found
;                  Not found : 0
; Remark ........: Case insensitive
; Author ........: Bert Kerkhof

Func aSearch(Const $aArray, Const $Value, Const $iColumn = 0, Const $iStart = 1, $iEnd = 0)
  If $iEnd = 0 Then $iEnd = $aArray[0]
  If _bounds3($aArray, $iColumn, $iStart, $iEnd) Then Return _errmess("aSearch", @error, @extended, 0)
  If $iColumn Then
    For $I = $iStart To $iEnd ; Case insensitive
      If ($aArray[$I])[$iColumn] = $Value Then Return $I
    Next
  Else
    For $I = $iStart To $iEnd ; Case insensitive
      If $aArray[$I] = $Value Then Return $I
    Next
  EndIf
  Return SetError(6, 0, 0)
EndFunc   ;==>aSearch

; #FUNCTION#
; Name ..........: aFindAll
; Description ...: Find multiple indexes of a search element in an array
; Syntax ........: aFindAll(Const $aArray, Const $Value[, $iStart = 1[, $iEnd = 0]])
; Parameters ....: $aArray ......... [const] an array of values.
;                  $Value .......... [const] The search value.
;                  $iStart, $iEnd .. [optional] range of array elements.
; Returns .......: Array of location index numbers
; Remark ........: Search is case insensitive
; Author ........: Bert Kerkhof

Func aFindAll(Const $aArray, Const $Value, Const $iStart = 1, $iEnd = 0)
  ; Collect multiple occurrences in an array
  ; Remark ........: Case insensitive
  ; Related .......: aStringFindAll aSearch

  If $iEnd = 0 Then $iEnd = $aArray[0]
  Local $aResult = aNew()
  If _bounds2($aArray, $iStart, $iEnd) Then Return _errmess("aFindAll", @error, 0, $aResult)
  For $I = $iStart To $iEnd
    If $aArray[$I] = $Value Then aAdd($aResult, $I)
  Next
  Return $aResult
EndFunc   ;==>aFindAll

; #FUNCTION#
; Name ..........: aStringLocate
; Description ...: Collect locations of multiple search matches
; Syntax ........: aStringLocate(Const $sString, Const $sSearch[, $iStart = 1[, $iEnd = 0]])
; Parameters ....: $sString .. [const] a string value.
;                  $sSearch .. [const] a string value.
;                  $iStart ... [optional] an integer value. Default is 1.
;                  $iEnd ..... [optional] an integer value. Default is 0.
; Returns .......: Array of numbers each pointing to first character of
;                  a found match.
; Remark ........: Search is case insensitive
; Author ........: Bert Kerkhof

Func aStringLocate(Const $sString, Const $sSearch, Const $iStart = 1, $iEnd = 0)
  If $iEnd > StringLen($sString) Then $iEnd = StringLen($sString)
  Local $aResult = aNew()
  If $iStart <1 or $iEnd<=0 Then Return _errmess("aStringLocate", 7, 0 , $aResult)
  Local $iOffset = $iStart
  While True
    $iOffset = StringInStr($sString, $sSearch, $STR_NOCASESENSE, 1, $iOffset, $iEnd - $iStart + 1)
    If $iOffset = 0 Then ExitLoop
    aAdd($aResult, $iOffset)
    $iOffset += 1
  WEnd
  Return $aResult
EndFunc   ;==>aStringLocate


; Array modification ==================================================

; #FUNCTION#
; Name ..........: aDelete
; Description ...: Delete one or more elements in an array
; Syntax ........: aDelete(Byref $aArray[, $iPos = 1[, $N = 1]])
; Parameters ....: $aArray .. Array to modify
;                  $iPos .... Element to delete.
;                             By default the first
;                  $N ....... Number of elements to delete.
;                             By default one
; Returns .......: Index of last item, equals the new length
; Comment .......: Deleting elements one by one is a slow process.
;                  Try to delete multiple elements at once top win
;                  execution speed.
; Author ........: Bert Kerkhof

Func aDelete(ByRef $aArray, Const $iPos = 1, $N = 1)
  If _bounds1($aArray) Then Return _errmess("aDelete", @error, 0, $aArray)
  $N = Min($N, $aArray[0] - $iPos + 1)
  If $iPos < 1 Or $N < 1 Then Return
  Local $I = $aArray[0] - $N
  For $J = $iPos To $I
    $aArray[$J] = $aArray[$J + $N]
  Next
  $aArray[0] = $I ; Update number of last index in element zero
  Return $aArray[0]
EndFunc   ;==>aDelete

; #FUNCTION#
; Name ..........: aInsert
; Description ...: Insert one element or an array of elements in an array
; Syntax ........: aInsert(Byref $aArray[, $iPos = 1[, $Value = 0]])
; Parameters ....: $aArray .. Array to modify
;                  $iPos .... Element number to insert row(s)
;                    by default insertion is in the first element
;                    Insertion possible up to LastIndex($aArray) +1
;                  $Value ... Value to insert.
;                    If Value is an array, all its elements are inserted.
;                    This causes the length of the receiving array to
;                    increase with the number of elements.
;                    If Value is a single element, only that one is
;                    is inserted.
; Returns .......: The new array length.
;                  The resulting array is passed by reference.
; Comment .......: Inserting elements one by one is a slow process.
;                  Try to insert multiple elements at once to win
;                  execution speed.
; Author ........: Bert Kerkhof

Func aInsert(ByRef $aArray, $iPos = 1, Const $Value = 0)
  If _bounds1($aArray) Then Return _errmess("aInsert", @error, 0, $aArray)
  If $iPos < 1 Then $iPos = 1
  If $iPos > $aArray[0] + 1 Then $iPos = $aArray[0] + 1

  Local $vPos
  Local $nSize = IsArray($Value) ? $Value[0] : 1
  $aArray[0] += $nSize ; Increase length of the array
  _NewDim($aArray)
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

; #FUNCTION#
; Name ..........: aFlat
; Description ...: Serially concatenates up to 16 array's / numbers / strings
; Syntax ........: aFlat([$A = 0[, $B = 0[, $C = 0[, $D = 0[, $E = 0[, $F = 0[, $G = 0[, $H = 0[, $I = 0[, $J = 0[,
;                  $K = 0[, $M = 0[, $N = 0[, $P = 0[, $Q = 0[, $R = 0]]]]]]]]]]]]]]]])
; Parameters ....: $A to $R ... [optional] mixture of single values and
;                               arrays.
; Returns .......: One-dimensional array of values
; Author ........: Bert Kerkhof

Func aFlat($A = 0, $B = 0, $C = 0, $D = 0, $E = 0, $F = 0, $G = 0, $H = 0, $I = 0, $J = 0, $K = 0, $M = 0, $N = 0, $P = 0, $Q = 0, $R = 0)
  Local $L, $aResult = aNew()
  Local $aArrayB = [@NumParams, $A, $B, $C, $D, $E, $F, $G, $H, $I, $J, $K, $M, $N, $P, $Q, $R]
  For $L = 1 To $aArrayB[0]
    aInsert($aResult, $aResult[0] + 1, $aArrayB[$L]) ; Insertion point is Lastindex($aArray)+1
  Next
  Return $aResult
EndFunc   ;==>aFlat

; #FUNCTION#
; Name ..........: aFill
; Description ...: Fill an array with a value
; Syntax ........: aFill(Byref $aArray[, $vValue = 0])
; Parameters ....: $aArray ... Array to be modified
;                  $vValue ... [optional] value to fill. If omitted,
;                              zero is assumed
; Returns .......: Number of updated array elements.
;                  The filled array is passed by reference.
; Comment .......: All elements are filled, even if aArray has multiple
;                  dimensions
; Author ........: Bert Kerkhof

Func aFill(ByRef $aArray, Const $vValue = 0)
  If _bounds1($aArray) Then Return _errmess("aFill", @error, 0, 0)
  Local $Row, $N = 0
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


; Sort ================================================================

; #FUNCTION#
; Name ..........: NumberCompare
; Description ...: Compare two numbers.
;                  Helper function for CombSort.
; Syntax ........: NumberCompare(Const $N1, Const $N2)
; Parameters ....: $N1, $N2 ... [const] numbers to compare
; Returns .......: Return 1, 0 or -1 depending on $N1 greater equal or
;                  less than $N2
; Also see ......: StringCompare (AutoIT function)
; Author ........: Bert Kerkhof

Func NumberCompare(Const $N1, Const $N2)
  If $N1 - $N2 > 0 Then Return 1 ; $N1 > $N2
  If $N2 - $N1 > 0 Then Return -1 ; $N1 < $N2
  Return 0
EndFunc   ;==>NumberCompare

; #FUNCTION#
; Name ..........: Lif
; Description ...: Select value depending on condition.
;                  Helper function for CombSort.
; Syntax ........: Lif(Const $Logic, Const $P1[, $P2 = ''])
; Parameters ....: $Logic ..... [const] condition.
;                  $P1 ........ [const] chosen if condition is True
;                  $P2 ........ [optional] if condition is False.
;                               Default is ''
; Returns .......: Value
; Remark ........: Warning: Although logic only chooses one, the function
;                  will activate both assignment choices. Use the 'ternari'
;                  syntax to avoid this.
; Author ........: Bert Kerkhof

Func Lif(Const $Logic, Const $P1, Const $P2 = '') ; Select value depending on condition
  Return $Logic ? $P1 : $P2
EndFunc   ;==>Lif

; #FUNCTION#
; Name ..........: aCombSort
; Description ...: Sort a two-dimensional array on one column
; Syntax ........: aCombSort(Byref $aArray[, $iColumn = 1[, $Ldescending = False[, $Lnumeric = False[, $iStart = 1[, $iEnd = 0]]]]])
; Parameters ....: $aArray ........ [in/out] a two dimensional array.
;                  $iColumn ....... Array column number that is the
;                                   search key
;                  $Ldescending ... [optional] Sort order flag.
;                                     False .. low values sorted first
;                                              (default)
;                                     True ... high values sorted first
;                  $Lnumeric ...... [optional] Value-type:
;                                     False .. sort key is string (default)
;                                     True ... sort key is numeric
;                  $iStart, iEnd .. [optional] Range of elements to be
;                                   sorted.
; Returns .......: Number of sort passes (for the curious)
;                  The sorted array is passed by reference.
; Author ........: Bert Kerkhof

Func aCombSort(ByRef $aArray, Const $iColumn = 1, Const $Ldescending = False, Const $Lnumeric = False, Const $iStart = 1, $iEnd = 0)
  If $iEnd = 0 Then $iEnd = $aArray[0]
  If _bounds3($aArray, $iColumn, $iStart, $iEnd) Then Return _errmess("aCombSort", @error, @extended, 0)
  If $iStart > $iEnd Then Return 0 ; Zero elements in array

  Local $Gap = $iEnd - $iStart ; Initialize gap size
  Local $nPass = 0, $R1, $R2
  Local $CompareLogic = $Ldescending ? -1 : 1
  Do
    $nPass += 1
    $Gap = Max(1, Floor($Gap / 1.3)) ; Apply the gap shrink factor
    Local $Count = $iStart, $Virgin = ($Gap = 1)
    If $Lnumeric Then ; Sort key is type numeric
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

; #FUNCTION#
; Name ..........: aSimpleCombSort
; Description ...: Perform aCombsort on a one-dimensional array
; Syntax ........: aSimpleCombSort(Byref $aArray[, $Ldescending = False[, $Lnumeric = False[, $iStart = 1[, $iEnd = 0]]]])
; Parameters ....: $aArray ....... [in/out] One-dimensional array to sort
;                  $Ldescending .. [optional] Sort order flag:
;                                    False .. low values sorted first
;                                             (default)
;                                    True ... high values sorted first
;                  $Lnumeric ..... [optional] Value-type:
;                                    False .. sort key is string (default)
;                                    True ... sort key is numeric
;                  iStart, iEnd .. [optional] Range of elements to be
;                                  sorted
; Returns .......: Number of sort passes (for the curious)
;                  The sorted array is passed by reference.
; Author ........: Bert Kerkhof

Func aSimpleCombSort(ByRef $aArray, Const $Ldescending = False, Const $Lnumeric = False, Const $iStart = 1, $iEnd = 0)
  If $iEnd = 0 Then $iEnd = $aArray[0]
  If _bounds2($aArray, $iStart, $iEnd) Then Return _errmess("aSimpleCombSort", @error, 0, 0)
  For $I = $iStart To $iEnd ; Pack :
    $aArray[$I] = Array($aArray[$I])
  Next
  Local $nPass = aCombSort($aArray, 1, $Ldescending, $Lnumeric, $iStart, $iEnd)
  For $I = $iStart To $iEnd ; UnPack :
    $aArray[$I] = ($aArray[$I])[1]
  Next
  Return $nPass
EndFunc   ;==>aSimpleCombSort


; Array type conversion ===============================================

; #FUNCTION#
; Name ..........: aDroste
; Description ...: Insert a count element at the zero element of an array.
; Syntax ........: aDroste(Const $aArray)
; Parameter......: $aArray .. 1Dim array to be converted
; Return value...: Drosted array
; Remarks........: Execution time 135 milliseconds on quad computer for
;                  100.000 elements array
; Author ........: Bert Kerkhof

Func aDroste(Const $aArray) ; Convert 1dim array to type Droste
  If Not IsArray($aArray) Then Return _errmess("aDroste", 1, 0, aNew())
  Local $nRow = UBound($aArray), $aResult = aNew($nRow)
  $aResult[0] = $nRow
  For $I = 1 To $nRow
    $aResult[$I] = $aArray[$I - 1]
  Next
  Return $aResult
EndFunc   ;==>aDroste


; Internals ===========================================================

; #FUNCTION#
; Name ..........: _NewDim
; Description ...: Internal allocation of array memory.
; Syntax ........: _NewDim(Byref $aArray)
; Parameters ....: $aArray .. [in/out] array to [re-]dimension
; Returns .......: None
; Author ........: Bert Kerkhof

Func _NewDim(ByRef $aArray)
  ; Stepwise array memory allocation (speeds up aAdd())
  If $aArray[0] < UBound($aArray) Then Return
  ; 16, 32, 64, 128, 256, 512, 1024, 2048 ..
  ReDim $aArray[2 ^ Ceiling(Log(Max($aArray[0] + 1, 16)) / Log(2))]
EndFunc   ;==>_NewDim

Func _errmess(Const $sOrigin, Const $error, Const $extended, Const $Rvalue)
  If @Compiled Then Return
  Local $aErr = aNew(7)
  $aErr[1] = 'Not an array'
  $aErr[2] = 'Not an array in second dim'
  $aErr[3] = 'An array in second dim is too short'
  $aErr[4] = 'More than two dimensions'
  $aErr[5] = 'Invalid array index'
  $aErr[6] = 'Invalid column number'
  $aErr[7] = 'Invalid string index'
  Local $sErr = ($error > 0 And $error < 6) ? @CRLF & $aErr[$error] : ""
  $sErr &= $extended ? @CRLF & "Element number " & $extended : ""
  MsgBox(64, @ScriptName, 'Error in function ' & $sOrigin & $sErr)
  Exit
  Return SetError($error, $extended, $Rvalue)
EndFunc   ;==>_errmess

Func _bounds1(Const $aArray)
  If @Compiled Then Return False
  If Not IsArray($aArray) Then Return SetError(1)
  Return False
EndFunc   ;==>_bounds1

Func _bounds2(Const $aArray, Const $iStart, Const $iEnd)
  If @Compiled Then Return False
  If Not IsArray($aArray) Then Return SetError(1)
  If $iStart < 1 Or $iStart>$aArray[0] Then Return SetError(5)
  If $iEnd <= 0 Or $iEnd > $aArray[0] Then Return SetError(5)
  Return False
EndFunc   ;==>_bounds2

Func _bounds3(Const $aArray, Const $iColumn, Const $iStart, Const $iEnd)
  If @Compiled Then Return False
  If _bounds2($aArray, $iStart, $iEnd) Then Return SetError(@error)
  If $iColumn < 0 Then Return SetError(6)
  If $iColumn = 0 Then Return False
  For $I = $iStart To $iEnd ; Check $iColumn for all rows:
    If Not IsArray($aArray[$I]) Then Return SetError(2, $I)
    If $iColumn > ($aArray[$I])[0] Then Return SetError(3, $I)
  Next
  Return False
EndFunc   ;==>_bounds3


; Tests ===============================================================

; #FUNCTION#
; Name ..........: UseDroste
; Description ...: Test the Ad, Search, CombSort and Recite functions
; Syntax ........: UseDroste()
; Parameters ....: None
; Returns .......: None
; Author ........: Bert Kerkhof

Func UseDroste()
  ; Build three dimensional Droste array:
  Local Enum $pNAME = 1, $pBIRTHDAY, $pOCCUPATION, $pMARRIED, $pPLACES
  aD("John de Vries", "1986-12-23", "Retail manager", True, aRecite("Bremen|Zwolle"))
  aD("Christiane Delore", "1982-05-01", "Teacher", True, aRecite("New York|Miami|Emmen"))
  aD("Marijke van Dongen", "unknown", "Carpenter", False, aRecite("Utrecht|London"))
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
  SetA($aPeople[aSearch($aPeople, 'Marijke van Dongen', $pNAME)], $pBIRTHDAY, "1983-04-25")

  ; Sort ascending on birthday:
  aCombSort($aPeople, $pBIRTHDAY, True)
  MsgBox(64, "From young to old", sRecite(aColumn($aPeople, $pNAME), ", "))

  ; More sort tests :
  Local $aR1 = Array(6, 5, 5, 5, 7)
  Local $nTest1 = aSimpleCombSort($aR1, True, True)
  MsgBox(64, "aSimpleCombSort. Descending numbers. nPass=" & $nTest1, sRecite($aR1))
  Local $nTest2 = aSimpleCombSort($aR1, False, True)
  MsgBox(64, "aSimpleCombSort. Ascending numbers. nPass=" & $nTest2, sRecite($aR1))
  Local $aR2 = aRecite("Portuguese|French|French|French|Spannish")
  Local $nTest3 = aSimpleCombSort($aR2, True, False)
  MsgBox(64, "aSimpleCombSort. Descending strings. nPass=" & $nTest3, sRecite($aR2))
  Local $nTest4 = aSimpleCombSort($aR2, False, False)
  MsgBox(64, "aSimpleCombSort. Ascending strings. nPass=" & $nTest4, sRecite($aR2))
  Local $aR0 = aNew()
  Local $nTest0 = aSimpleCombSort($aR0, True, True)
  MsgBox(64, "aSimpleCombSort. Zero elements. nPass=" & $nTest0, sRecite($aR0))

  ; aStringFindAll :
  Local $S = "De kapper kapt knap maar de knecht van de knappe kapper kapt knapper dan de knappe kapper kappen kan"
  MsgBox(64, "Positions in string of 'kapper'", sRecite(aStringLocate($S, "kapper"), ", "))

EndFunc   ;==>UseDroste
; UseDroste()

; #FUNCTION#
; Name ..........: sCountWord
; Description ...: Literal for rank number
; Syntax ........: sCountWord(Const $N[, $Luppercase = False])
; Parameters ....: $N ........... [const] rank number
;                  $Luppercase .. [optional] case of the first character of
;                                 the word.
; Returns .......: Literal string
; Author ........: Bert Kerkhof

Func sCountWord(Const $N, Const $Luppercase = False)
  Local $aCountW = aRecite("First|Second|Third|Fourth|Fifth|Sixth|Seventh|Eight|Nineth|Tenth")
  Local $sR = ($N > 10) ? $N & 'th' : $aCountW[$N]
  Return $Luppercase ? $sR : StringLower($sR)
EndFunc   ;==>CountWord

Func TestsCountWord()
  MsgBox(64, "TestsCountWord", sCountWord(2))
EndFunc
; sTestsCountWord()

; #FUNCTION#
; Name ..........: Test_NewDim
; Description ...: Demonstrate the dynamic allocation of array memory
; Syntax ........: Test_NewDim()
; Parameters ....: None
; Returns .......: None
; Author ........: Bert Kerkhof

Func Test_NewDim()
  Local $aArray = aNew()
  Local $cOut = 'aNew, inital dimension: ' & UBound($aArray) & @LF
  aInsert($aArray, 1, aNew(15))
  $cOut &= 'after adding 15 values: ' & UBound($aArray) & @LF
  aAdd($aArray, 1)
  $cOut &= 'after adding the 16th value: ' & UBound($aArray)
  MsgBox(64, "Test_NewDim", $cOut)
EndFunc   ;==>Test_NewDim
; Test_NewDim()

; End =================================================================
