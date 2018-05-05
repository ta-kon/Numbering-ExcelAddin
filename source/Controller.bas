Attribute VB_Name = "Controller"

' 連番太郎
' 
' Copyright (c) 2018 ta-kon
'
' The MIT License (MIT)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit

Public Function getDateTime() As String
  getDateTime = Format(Now, FORMAT_DATE)
End Function

Public Function shapeNumbering(Optional optStartNumber As Long = -1) As Long
  Const processName As String = "採番処理"

  Call startProcess(processName)

  If (Not(isSelectedShapes())) Then
    Call msgNotSelectedShape(processName)
    Exit Function
  End If

  Dim shapeArray() As Shape
  shapeArray = getShapeArray()

  If (SELECT_SORT_ORDER <> SORT_ORDER_SELECT) Then
    ' 選択順でない場合のみソート
    Call sort(shapeArray)
  End If

  Dim startNumber As Long

  If (optStartNumber <> -1) Then
    ' 採番の初期値の指定がある場合 (図形の複製で使用)
    startNumber = optStartNumber
  Else
    ' 採番の初期値がない場合
    ' 最初の図形または、開始値の読み込み
    startNumber = getStartNumber(shapeArray)
  End If

  Dim lastNumber As Long
  lastNumber = setShapeNumbering(shapeArray, startNumber)

  Dim processDetail As String
  processDetail = "開始値:" & startNumber & ", 終了値:" & lastNumber
  Call endProcess(processName, getShapeArrayCount(), processDetail)

  shapeNumbering = lastNumber
End Function

Public Sub shapeClear()
  Const processName As String = "テキストの除去"

  Call startProcess(processName)

  If (Not(isSelectedShapes())) Then
    Call msgNotSelectedShape(processName)
    Exit Sub
  End If

  Dim shapeArray() As Shape
  shapeArray = getShapeArray()

  Call setShapeClear(shapeArray)
  Call endProcess(processName, getShapeArrayCount())
End Sub

Public Sub shapeClone()
  Const processName As String = "図形の複製"

  Call startProcess(processName)

  If (Not(isSelectedShapes())) Then
    Call msgNotSelectedShape(processName)
    Exit Sub
  End If

  If (Not(isSingleSelectedShape())) Then
    Call msgNotSingleSelectedShape(processName)
    Exit Sub
  End If

  Dim shapeArray() As Shape
  shapeArray = getShapeArray()

  Dim startNumber As Long
  ' 複製前の最初の値を取得
  startNumber = getStartNumber(shapeArray)

  Dim cloneCount As Long
  cloneCount = Application.InputBox(Prompt:="図形を複製する数を入力してください。", Title:="図形の複製", Type:=1)

  If (cloneCount <=0) Then
    Call setStatusBar(processName & "を行うためには、正数を入力してください。")
    Exit Sub
  End If

  Call cloneShape(shapeArray(0), cloneCount)

  ' テキストの色が変わるかもしれないため、テキストを削除
  Call shapeClear()

  Dim lastNumber As Long
  ' 複製後の図形を最初の値から採番
  lastNumber = shapeNumbering(startNumber)

  Dim processDetail As String
  processDetail = "開始値:" & startNumber & ", 終了値:" & lastNumber
  Call endProcess(processName, getShapeArrayCount(), processDetail)
End Sub

Public Sub autoSelect()
  Const processName As String = "自動選択"

  Dim selectCout As Long
  If (Not(IS_SELECT_COLOR Or IS_SELECT_FIGURE Or IS_SELECT_SIZE)) Then
    ActiveSheet.Shapes.SelectAll
    selectCout = getShapeArrayCount()
  Else
    If (Not(isSelectedShapes())) Then
      Call msgNotSelectedShape(processName)
      Exit Sub
    End If

    If (Not(isSingleSelectedShape())) Then
      Call msgNotSingleSelectedShape(processName)
      Exit Sub
    End If

    selectCout = autoEqualsSelectShape()
  End If

  Call endProcess(processName, selectCout)
End Sub

Private Sub cloneShape(ByVal baseShape As Shape, ByVal cloneCount As Long)

  Dim shapeTop As Long
  shapeTop = baseShape.Top

  Dim count As Long
  For count = 1 To cloneCount
    shapeTop = shapeTop + (baseShape.Height + COLONE_MARGINE)

    ' 図形を複製
    With baseShape.Duplicate
      .Top  = shapeTop
      .Left = baseShape.Left
      .Select Replace:=FALSE
    End With
  Next count
End Sub

Private Function autoEqualsSelectShape() As Long
  Dim shapeArray() As Shape
  shapeArray = getShapeArray()

  Dim searchShape As Shape
  Set searchShape = shapeArray(0)

  Dim selectCout As Long
  selectCout = 0

  Dim shape As Shape
  For Each shape In ActiveSheet.Shapes
    If (equalsShape(searchShape, shape)) Then
      shape.Select Replace:=FALSE
      selectCout = selectCout + 1
    End If
  Next shape

  autoEqualsSelectShape = selectCout
End Function

Private Function equalsShape(ByVal leftShape As Shape, ByVal rightShape As Shape) As Boolean

  ' 独立性が高い順から比較するため、色の情報から比較
  ' VBAでは短絡評価をしないので、一つひとつ評価
  If (IS_SELECT_COLOR) Then
    If (leftShape.Fill.ForeColor.RGB <> rightShape.Fill.ForeColor.RGB) Then
      equalsShape = FALSE
      Exit Function
    End If
  End If

  If (IS_SELECT_FIGURE) Then
    If (leftShape.AutoShapeType <> rightShape.AutoShapeType) Then
      equalsShape = FALSE
      Exit Function
    End If
  End If

  If (IS_SELECT_SIZE) Then
    If (leftShape.Width <> rightShape.Width) Then
      equalsShape = FALSE
      Exit Function
    End If

    If (leftShape.Height <> rightShape.Height) Then
      equalsShape = FALSE
      Exit Function
    End If
  End If

  equalsShape = TRUE
End Function

Private Function isSelectedShapes() As Boolean
  ' オートシェイプを選択しているか
  isSelectedShapes = (getShapeArrayCount() > 0)
End Function

Private Function isSingleSelectedShape() As Boolean
  isSingleSelectedShape = (getShapeArrayCount() = 1)
End Function

Private Function getShapeArrayCount() As Integer
  ' 選択中のオートシェイプ数
  getShapeArrayCount = UBound(getShapeArray())
End Function

Private Sub setShapeClear(ByRef shapeArray() As Shape)
  ' オートシェイプのテキストを除去

  Dim shape As Shape
  Dim index As Long
  For index = LBound(shapeArray) To UBound(shapeArray) - 1
    Set shape = shapeArray(index)
    Call setShapeText(shape, "")
  Next index
End Sub

Private Function getStartNumber(ByRef shapeArray() As Shape) As Long
  ' 採番を行う最初の数値を取得

  If (ENABLED_SART_NUMBER) Then
    ' 開始値の指定が有るとき
    getStartNumber = START_NUM
    Exit Function
  End If

  ' 開始値の指定が無いとき
  Dim index As Long
  index = LBound(shapeArray)

  Dim shape As Shape
  Set shape = shapeArray(index)

  ' 入力した開始値で採番します。
  On Error GoTo ERR_NUMBER_FORMAT
    getStartNumber = CLng(getShapeText(shape))
  Exit Function

  ERR_NUMBER_FORMAT:
    getStartNumber = 1
End Function

Private Function setShapeNumbering(ByRef shapeArray() As Shape, Optional startNumber As Long = 1) As Long
  ' 図形に番号を付与

  ' 図形に付与する数値
  Dim number As Long
  number = startNumber

  ' 図形の配列番号
  Dim index As Long
  For index = LBound(shapeArray) To UBound(shapeArray) - 1

    Dim shape As Shape
    Set shape = shapeArray(index)

    If (IS_CHANGE_TEXT) Then
      Call changeColor(shape, Str(number))
    End If

    Call setShapeText(shape, number)

    ' 最後に採番した番号
    setShapeNumbering = number

    ' 図形に付与する番号を増やす
    number = number + 1
  Next index

End Function

Private Sub changeColor(ByVal shape As Shape, ByVal compareText As String)
    Dim shapeText As String
    shapeText = getShapeText(shape)

    if (shapeText <> "" And shapeText <> compareText) Then
      Call setShapeTextColor(shape)
    End If
End Sub

Private Function isDescShape(ByRef leftShape As Shape, ByRef rightShape As Shape) As Boolean

  Select Case SELECT_SORT_ORDER
    Case SORT_ORDER_ROW
      isDescShape = isDescShapeRow(leftShape, rightShape)
      Exit Function

    Case SORT_ORDER_COLUM
      isDescShape = isDescShapeColum(leftShape, rightShape)
      Exit Function

    Case SORT_ORDER_SELECT
      ' ソート対象ではない
      isDescShape = FALSE
      Exit Function
  End Select

End Function

Private Function isDescShapeRow(ByRef leftShape As Shape, ByRef rightShape As Shape) As Boolean

  ' 同じ高さ (行) に存在している場合
  If (isCollision(leftShape.Top, rightShape.Top, COLLISION)) Then
      ' 左からの座標を比較 (leftShapeの方が大きいとき)
      isDescShapeRow = isDescLocation(leftShape.Left, rightShape.Left)
    Exit Function
  End If

  ' 上からの座標を比較 (leftShapeの方が大きいとき)
  isDescShapeRow = isDescLocation(leftShape.Top, rightShape.Top)
End Function

Private Function isDescShapeColum(ByRef leftShape As Shape, ByRef rightShape As Shape) As Boolean

  ' 同じ列 に存在している場合
  If (isCollision(leftShape.Left, rightShape.Left, COLLISION)) Then
      ' 上からの座標を比較 (leftShapeの方が大きいとき)
      isDescShapeColum = isDescLocation(leftShape.Top, rightShape.Top)
    Exit Function
  End If

  ' 左からの座標を比較 (leftShapeの方が大きいとき)
  isDescShapeColum = isDescLocation(leftShape.Left, rightShape.Left)
End Function

Private Function isCollision(ByVal left As Long, ByVal right As Long, ByVal collision As Long) As Boolean
  ' 幅を考慮した数値の比較
  isCollision = (Abs(left - right) < collision)
End Function

Private Function isDescLocation(ByVal left As Long, ByVal right As Long) As Boolean
    isDescLocation = (left > right)
End Function

Private Sub sort(ByRef shapeArray() As Shape)
  ' 挿入ソート (1番に採番したい順に並び替え)
  ' ソート対象数は大抵50件以下であるため、挿入ソートを採用

  Dim tmpShape As Shape
  Dim index As Long
  For index = LBound(shapeArray) + 1 To UBound(shapeArray) - 1

    Set tmpShape = shapeArray(index)

    ' 前の要素のほうが大きい場合
    If (isDescShape(shapeArray(index - 1), tmpShape)) Then

      Dim beforeIndex As Long
      beforeIndex = index

      ' 前の要素を後ろに一つずつずらす
      Do
          ' 前の要素を後ろに移動
          Set shapeArray(beforeIndex) = shapeArray(beforeIndex - 1)

          beforeIndex = beforeIndex - 1

          ' これ以上前の要素がない場合
          If (beforeIndex = 0) Then
            Exit Do
          End If

      Loop While (isDescShape(shapeArray(beforeIndex - 1), tmpShape))

      ' データを挿入
      Set shapeArray(beforeIndex) = tmpShape
    End If

  Next index
End Sub