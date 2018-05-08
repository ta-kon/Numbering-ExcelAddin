Attribute VB_Name = "Model"

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

' �v���W�F�N�g��
Public Const PROJECT_NAME As String = "�A�ԑ��Y"

' �v���W�F�N�g�̃o�[�W����
Public Const PROJECT_REVISION As String = "20180428"

' �����̏o�͌`��
Public Const FORMAT_DATE As String = "yyyy/MM/dd hh:mm:ss"

' ���{��
Public I_RIBBON_UI As IRibbonUI

' ' �����񂾂Ɣ��肷�鍂��
' Public Const COLLISION_HEIGHT As Long = 20

' ' �����s���Ɣ��肷�镝
' Public Const COLLISION_WITH As Long = 40

' �ߐڕ�
Public COLLISION As Double

Public COLLISION_NUM_STRING As String

' �s�P�ʂō̔�
Public Const SORT_ORDER_ROW As Integer = 0

' ��P�ʂō̔�
Public Const SORT_ORDER_COLUM As Integer = 1

' �I���������ō̔�
Public Const SORT_ORDER_SELECT As Integer = 2

' �̔ԏ���
Public SELECT_SORT_ORDER As Integer

' �ύX�̂������e�L�X�g�̐F (�F�p���b�g ��:3)
Public Const CHANGE_TEXT_COLOR_INDEX As Integer = 3

' �ύX�̂������̔Ԃ͐ԐF�ɂ���
Public IS_CHANGE_TEXT As Boolean

' �J�n�l��L���ɂ���
Public ENABLED_SART_NUMBER As Boolean

' �J�n�l
Public START_NUM As Long

' �����`��I������
Public IS_SELECT_FIGURE As Boolean

' �����傫����I������
Public IS_SELECT_SIZE As Boolean

' �����F��I������
Public IS_SELECT_COLOR AS Boolean

' �}�`�̕�������margin
Public Const COLONE_MARGINE As Long = 2

' ������
Public Sub initModel()
  ' ���{���N�����ɏ�����

  SELECT_SORT_ORDER = SORT_ORDER_ROW
  COLLISION_NUM_STRING = "0.50�{ (�W��)"
  COLLISION = 0.50
  START_NUM = 1
  IS_CHANGE_TEXT   = TRUE
  IS_SELECT_FIGURE = TRUE
  IS_SELECT_COLOR  = TRUE
  IS_SELECT_SIZE   = TRUE
  ENABLED_SART_NUMBER = FALSE
End Sub
