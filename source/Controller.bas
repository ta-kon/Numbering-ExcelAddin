Attribute VB_Name = "Controller"

' �A�ԑ��Y
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
  Const processName As String = "�̔ԏ���"

  Call startProcess(processName)

  If (Not(isSelectedShapes())) Then
    Call msgNotSelectedShape(processName)
    Exit Function
  End If

  Dim shapeArray() As Shape
  shapeArray = getShapeArray()

  If (SELECT_SORT_ORDER <> SORT_ORDER_SELECT) Then
    ' �I�����łȂ��ꍇ�̂݃\�[�g
    Call sort(shapeArray)
  End If

  Dim startNumber As Long

  If (optStartNumber <> -1) Then
    ' �̔Ԃ̏����l�̎w�肪����ꍇ (�}�`�̕����Ŏg�p)
    startNumber = optStartNumber
  Else
    ' �̔Ԃ̏����l���Ȃ��ꍇ
    ' �ŏ��̐}�`�܂��́A�J�n�l�̓ǂݍ���
    startNumber = getStartNumber(shapeArray)
  End If

  Dim lastNumber As Long
  lastNumber = setShapeNumbering(shapeArray, startNumber)

  Dim processDetail As String
  processDetail = "�J�n�l:" & startNumber & ", �I���l:" & lastNumber
  Call endProcess(processName, getShapeArrayCount(), processDetail)

  shapeNumbering = lastNumber
End Function

Public Sub shapeClear()
  Const processName As String = "�e�L�X�g�̏���"

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
  Const processName As String = "�}�`�̕���"

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
  ' �����O�̍ŏ��̒l���擾
  startNumber = getStartNumber(shapeArray)

  Dim cloneCount As Long
  cloneCount = Application.InputBox(Prompt:="�}�`�𕡐����鐔����͂��Ă��������B", Title:="�}�`�̕���", Type:=1)

  If (cloneCount <=0) Then
    Call setStatusBar(processName & "���s�����߂ɂ́A��������͂��Ă��������B")
    Exit Sub
  End If

  Call cloneShape(shapeArray(0), cloneCount)

  ' �e�L�X�g�̐F���ς�邩������Ȃ����߁A�e�L�X�g���폜
  Call shapeClear()

  Dim lastNumber As Long
  ' ������̐}�`���ŏ��̒l����̔�
  lastNumber = shapeNumbering(startNumber)

  Dim processDetail As String
  processDetail = "�J�n�l:" & startNumber & ", �I���l:" & lastNumber
  Call endProcess(processName, getShapeArrayCount(), processDetail)
End Sub

Public Sub autoSelect()
  Const processName As String = "�����I��"

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

    ' �}�`�𕡐�
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

  ' �Ɨ����������������r���邽�߁A�F�̏�񂩂��r
  ' VBA�ł͒Z���]�������Ȃ��̂ŁA��ЂƂ]��
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
  ' �I�[�g�V�F�C�v��I�����Ă��邩
  isSelectedShapes = (getShapeArrayCount() > 0)
End Function

Private Function isSingleSelectedShape() As Boolean
  isSingleSelectedShape = (getShapeArrayCount() = 1)
End Function

Private Function getShapeArrayCount() As Integer
  ' �I�𒆂̃I�[�g�V�F�C�v��
  getShapeArrayCount = UBound(getShapeArray())
End Function

Private Sub setShapeClear(ByRef shapeArray() As Shape)
  ' �I�[�g�V�F�C�v�̃e�L�X�g������

  Dim shape As Shape
  Dim index As Long
  For index = LBound(shapeArray) To UBound(shapeArray) - 1
    Set shape = shapeArray(index)
    Call setShapeText(shape, "")
  Next index
End Sub

Private Function getStartNumber(ByRef shapeArray() As Shape) As Long
  ' �̔Ԃ��s���ŏ��̐��l���擾

  If (ENABLED_SART_NUMBER) Then
    ' �J�n�l�̎w�肪�L��Ƃ�
    getStartNumber = START_NUM
    Exit Function
  End If

  ' �J�n�l�̎w�肪�����Ƃ�
  Dim index As Long
  index = LBound(shapeArray)

  Dim shape As Shape
  Set shape = shapeArray(index)

  ' ���͂����J�n�l�ō̔Ԃ��܂��B
  On Error GoTo ERR_NUMBER_FORMAT
    getStartNumber = CLng(getShapeText(shape))
  Exit Function

  ERR_NUMBER_FORMAT:
    getStartNumber = 1
End Function

Private Function setShapeNumbering(ByRef shapeArray() As Shape, Optional startNumber As Long = 1) As Long
  ' �}�`�ɔԍ���t�^

  ' �}�`�ɕt�^���鐔�l
  Dim number As Long
  number = startNumber

  ' �}�`�̔z��ԍ�
  Dim index As Long
  For index = LBound(shapeArray) To UBound(shapeArray) - 1

    Dim shape As Shape
    Set shape = shapeArray(index)

    If (IS_CHANGE_TEXT) Then
      Call changeColor(shape, Str(number))
    End If

    Call setShapeText(shape, number)

    ' �Ō�ɍ̔Ԃ����ԍ�
    setShapeNumbering = number

    ' �}�`�ɕt�^����ԍ��𑝂₷
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
      ' �\�[�g�Ώۂł͂Ȃ�
      isDescShape = FALSE
      Exit Function
  End Select

End Function

Private Function isDescShapeRow(ByRef leftShape As Shape, ByRef rightShape As Shape) As Boolean

  ' �������� (�s) �ɑ��݂��Ă���ꍇ
  If (isCollision(leftShape.Top, rightShape.Top, COLLISION)) Then
      ' ������̍��W���r (leftShape�̕����傫���Ƃ�)
      isDescShapeRow = isDescLocation(leftShape.Left, rightShape.Left)
    Exit Function
  End If

  ' �ォ��̍��W���r (leftShape�̕����傫���Ƃ�)
  isDescShapeRow = isDescLocation(leftShape.Top, rightShape.Top)
End Function

Private Function isDescShapeColum(ByRef leftShape As Shape, ByRef rightShape As Shape) As Boolean

  ' ������ �ɑ��݂��Ă���ꍇ
  If (isCollision(leftShape.Left, rightShape.Left, COLLISION)) Then
      ' �ォ��̍��W���r (leftShape�̕����傫���Ƃ�)
      isDescShapeColum = isDescLocation(leftShape.Top, rightShape.Top)
    Exit Function
  End If

  ' ������̍��W���r (leftShape�̕����傫���Ƃ�)
  isDescShapeColum = isDescLocation(leftShape.Left, rightShape.Left)
End Function

Private Function isCollision(ByVal left As Long, ByVal right As Long, ByVal collision As Long) As Boolean
  ' �����l���������l�̔�r
  isCollision = (Abs(left - right) < collision)
End Function

Private Function isDescLocation(ByVal left As Long, ByVal right As Long) As Boolean
    isDescLocation = (left > right)
End Function

Private Sub sort(ByRef shapeArray() As Shape)
  ' �}���\�[�g (1�Ԃɍ̔Ԃ��������ɕ��ёւ�)
  ' �\�[�g�Ώې��͑��50���ȉ��ł��邽�߁A�}���\�[�g���̗p

  Dim tmpShape As Shape
  Dim index As Long
  For index = LBound(shapeArray) + 1 To UBound(shapeArray) - 1

    Set tmpShape = shapeArray(index)

    ' �O�̗v�f�̂ق����傫���ꍇ
    If (isDescShape(shapeArray(index - 1), tmpShape)) Then

      Dim beforeIndex As Long
      beforeIndex = index

      ' �O�̗v�f�����Ɉ�����炷
      Do
          ' �O�̗v�f�����Ɉړ�
          Set shapeArray(beforeIndex) = shapeArray(beforeIndex - 1)

          beforeIndex = beforeIndex - 1

          ' ����ȏ�O�̗v�f���Ȃ��ꍇ
          If (beforeIndex = 0) Then
            Exit Do
          End If

      Loop While (isDescShape(shapeArray(beforeIndex - 1), tmpShape))

      ' �f�[�^��}��
      Set shapeArray(beforeIndex) = tmpShape
    End If

  Next index
End Sub