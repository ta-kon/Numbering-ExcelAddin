Attribute VB_Name = "View"

Option Explicit

' ���{���̋@�\�ۑ��Ɏg�p
#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As LongPtr)
#Else
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
#End If

Sub onActionNoumberingButton(ByVal control As IRibbonControl)
    Call shapeNumbering()
End Sub

Sub onActionClearButton(ByVal control As IRibbonControl)
    Call shapeClear()
End Sub

Sub onActionCloneButton(ByVal control As IRibbonControl)
    Call shapeClone()
End Sub

Sub onActionAutoSelectButton(ByVal control As IRibbonControl)
    Call autoSelect()
End Sub

Public Sub setStatusBar(ByVal message As String)
  Application.StatusBar = message
End Sub

Public Sub showMsgBox(ByVal message As String)
  MsgBox message, vbOKOnly, PROJECT_NAME
End Sub

Public Sub setShapeText(ByRef shape As Shape, ByVal text As String)
  shape.TextFrame.Characters.Text = text
End Sub

Public Sub setShapeTextColor(ByRef shape)
  shape.TextFrame.Characters.Font.ColorIndex = CHANGE_TEXT_COLOR_INDEX
End Sub

Public Function getShapeText(ByRef shape As Shape) As String
  getShapeText = shape.TextFrame.Characters.Text
End Function

Public Sub setMsgAndStatus(ByVal message As String)
  setStatusBar(message)
  MsgBox message, vbOKOnly, PROJECT_NAME
End Sub

Public Sub msgNotSelectedShape(ByVal processName As String)
  Dim message As String
  message = processName & "���s���}�`��I�����Ă��������B"
  Call setMsgAndStatus(message)
End Sub

Public Sub msgNotSingleSelectedShape(ByVal processName As String)
  Dim message As String
  message = processName & "���s���}�`��1�����I�����Ă��������B"
  Call setMsgAndStatus(message)
End Sub

Public Sub startProcess(ByVal processName)
  setStatusBar(processName & "���J�n���܂��B")
End Sub

Public Sub endProcess(ByVal processName As String, ByVal count As Long, Optional detailMessage As String = "")

  Dim processMessage As String

  If (detailMessage <> "") Then
    processMessage = processName & "���������܂����B ����:" & count & "��, " & detailMessage & ", �����I������:"  & getDateTime()
  Else
    processMessage = processName & "���������܂����B ����:" & count & "��, �����I������:"  & getDateTime()
  End If

  setStatusBar(processMessage)
End Sub

Sub onRibbonLoad(ByRef ribbon As IRibbonUI)
  ' ���{���̏�������

  Call initModel()

  '���{���̃|�C���^�����W�X�g���ɋL�^
  SaveSetting "NumberingApp", "Main", "RibbonPointer", CStr(ObjPtr(ribbon))

  ' ���{���̕\�����X�V�ł���悤�ɂ��邽��
  Set I_RIBBON_UI = ribbon
  ' ���{�����X�V
  I_RIBBON_UI.Invalidate
End Sub

Sub getStartNumberEnabled(ByRef control As IRibbonControl, ByRef returnedVal)
  ' ���͂����J�n�l�ō̔Ԃ��܂��B
  returnedVal = ENABLED_SART_NUMBER
End Sub

Sub getStartNumberText(ByRef control As IRibbonControl, ByRef returnedVal)
  ' ���͂����J�n�l�ō̔Ԃ��܂��B
  ' If (ENABLED_SART_NUMBER) Then
  ' End If
  returnedVal = START_NUM
End Sub

Sub onChangeStartNumberText(ByRef control As IRibbonControl, ByRef text As String)
  ' ���͂����J�n�l�ō̔Ԃ��܂��B
  On Error GoTo ERR_NUMBER_FORMAT
    START_NUM = CLng(text)
  Exit Sub

  ERR_NUMBER_FORMAT:
    START_NUM = 1
    Const message As String = "�J�n�l�͐�������͂��Ă��������B"
    Call setMsgAndStatus(message)
End Sub

Sub getSortOrderSelectedIndex(ByRef control As IRibbonControl, ByRef index)
  ' �̔ԏ����F
  index = SELECT_SORT_ORDER
End Sub

Sub getCollisionText(ByRef control As IRibbonControl, ByRef returnedVal)
  ' �ߐڕ��F
  returnedVal = COLLISION_NUM_STRING
End Sub

Sub onCollisionTextChange(ByRef control As IRibbonControl, ByRef text As String)
  ' �ߐڕ��F
  COLLISION_NUM_STRING = text

  Select Case text
    Case "0�{ (��א�)"
      COLLISION = 0
    Case "0.25�{"
      COLLISION = 0.25
    Case "0.50�{ (�W��)"
      COLLISION = 0.50
    Case "0.75�{"
      COLLISION = 0.75
    Case "1.00�{"
      COLLISION = 1.00
    Case Else
      Dim strDbl As String
      strDbl = Trim(text)
      strDbl = Replace(strDbl, "�{", "")
      strDbl = Trim(strDbl)
      If IsNumeric(strDbl) Then
        COLLISION = Abs(CDbl(strDbl))
      Else
        Dim message As String
        message = "���l����͂��Ă��������B(�P��: �{), ���݂̒l=" & text
        Call setMsgAndStatus(message)
      End If
  End Select

End Sub

Sub getSelectFigureCheckBoxPressed(ByRef control As IRibbonControl, ByRef returnedVal)
  ' �����`��I������
  returnedVal = IS_SELECT_FIGURE
End Sub

Sub getSelectSizeCheckBoxPressed(ByRef control As IRibbonControl, ByRef returnedVal)
  ' �����傫����I������
  returnedVal = IS_SELECT_SIZE
End Sub

Sub getSelectColorCheckBoxPressed(ByRef control As IRibbonControl, ByRef returnedVal)
  ' �����F��I������
  returnedVal = IS_SELECT_COLOR
End Sub

Sub getChangeTextPressed(ByRef control As IRibbonControl, ByRef returnedVal)
  ' �ύX�̂������̔Ԃ͐ԐF�ɂ���
  returnedVal = IS_CHANGE_TEXT
End Sub

Sub onActionChangeText(ByRef control As IRibbonControl, ByRef pressed As Boolean)
  ' �ύX�̂������̔Ԃ͐ԐF�ɂ���
  IS_CHANGE_TEXT = pressed
End Sub

Sub getContinueNumberPressed(ByRef control As IRibbonControl, ByRef returnedVal)
  ' �ŏ��̐}�`�̐��l����̔Ԃ���
  returnedVal = NOT(ENABLED_SART_NUMBER)
End Sub

Sub onActionContinueNumberCheckBox(ByRef control As IRibbonControl, ByRef pressed As Boolean)
  ' �ŏ��̐}�`�̐��l����̔Ԃ���
  ENABLED_SART_NUMBER = NOT(pressed)

  If I_RIBBON_UI Is Nothing Then
    #If VBA7 And Win64 Then
      Set I_RIBBON_UI = GetRibbon(CLngPtr(GetSetting("NumberingApp", "Main", "RibbonPointer")))
    #Else
      Set I_RIBBON_UI = GetRibbon(CLng(GetSetting("NumberingApp", "Main", "RibbonPointer")))
    #End If
  End If

  I_RIBBON_UI.InvalidateControl "startNumberEditBox"
End Sub

Sub onActionSelectFigureCheckBox(ByRef control As IRibbonControl, ByRef pressed As Boolean)
  IS_SELECT_FIGURE = pressed
End Sub

Sub onActionSelectSizeCheckBox(ByRef control As IRibbonControl, ByRef pressed As Boolean)
  IS_SELECT_SIZE = pressed
End Sub

Sub onActionSelectColorCheckBox(ByRef control As IRibbonControl, ByRef pressed As Boolean)
  IS_SELECT_COLOR = pressed
End Sub

' �ȉ���Function GetRibbon�́A�ȉ��̃T�C�g���Q�l�ɂȂ�܂����B
' [IRibbonUI�I�u�W�F�N�g��Nothing�ɂȂ����Ƃ��̑Ώ��@](http://www.ka-net.org/ribbon/ri64.html)
#If VBA7 And Win64 Then
Private Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
  Dim p As LongPtr
#Else
Private Function GetRibbon(ByVal lRibbonPointer As Long) As Object
  Dim p As Long
#End If
  Dim ribbonObj As Object

  MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = ribbonObj
  p = 0: MoveMemory ribbonObj, p, LenB(p)
End Function

Sub onActionSortOrderSlected(ByRef control As IRibbonControl, ByRef itemID As String, ByRef index As Integer)
  ' �̔Ԃ��鏇����ݒ肵�܂��B
  Select Case itemID
    Case "ROW_SORT"
      SELECT_SORT_ORDER = SORT_ORDER_ROW
    Case "COLUMN_SORT"
      SELECT_SORT_ORDER = SORT_ORDER_COLUM
    Case "SELECT_SORT"
      SELECT_SORT_ORDER = SORT_ORDER_SELECT
  End Select
End Sub

Public Function getShapeArray As Shape()

  On Error GoTo ERR_NOT_FOUND_SHAPE
  ' �}�`��I�����Ă��Ȃ��Ƃ��� ERR_NOT_FOUND_SHAPE �֑J��

    Dim shapeRange As ShapeRange
    Set shapeRange = getShapeRange()

    Dim shapeRangeCout As Long
    shapeRangeCout = shapeRange.Count

    Dim shapeArray() As Shape
    ReDim shapeArray(shapeRangeCout)

    Dim index As Integer
    index = 0
    Dim shape As Shape
    For Each shape In shapeRange
      Set shapeArray(index) = shape
      index = index + 1
    Next shape

    getShapeArray = shapeArray
  Exit Function

  ERR_NOT_FOUND_SHAPE:
  ' �}�`��I�����Ă��Ȃ��Ƃ�
    Dim shapeEmptyArray(0) As Shape
    getShapeArray = shapeEmptyArray
End Function

Private Function getShapeRange() As ShapeRange
  ' �}�`��I�����������Ŏ擾
  ' �}�`���擾�ł��Ȃ������ꍇ��Error
    Set getShapeRange = Selection.ShapeRange

End Function