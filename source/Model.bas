Attribute VB_Name = "Model"

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

' プロジェクト名
Public Const PROJECT_NAME As String = "連番太郎"

' プロジェクトのバージョン
Public Const PROJECT_REVISION As String = "20180428"

' 日時の出力形式
Public Const FORMAT_DATE As String = "yyyy/MM/dd hh:mm:ss"

' リボン
Public I_RIBBON_UI As IRibbonUI

' ' 同じ列だと判定する高さ
' Public Const COLLISION_HEIGHT As Long = 20

' ' 同じ行だと判定する幅
' Public Const COLLISION_WITH As Long = 40

' 近接幅
Public COLLISION As Double

Public COLLISION_NUM_STRING As String

' 行単位で採番
Public Const SORT_ORDER_ROW As Integer = 0

' 列単位で採番
Public Const SORT_ORDER_COLUM As Integer = 1

' 選択した順で採番
Public Const SORT_ORDER_SELECT As Integer = 2

' 採番順序
Public SELECT_SORT_ORDER As Integer

' 変更のあったテキストの色 (色パレット 赤:3)
Public Const CHANGE_TEXT_COLOR_INDEX As Integer = 3

' 変更のあった採番は赤色にする
Public IS_CHANGE_TEXT As Boolean

' 開始値を有効にする
Public ENABLED_SART_NUMBER As Boolean

' 開始値
Public START_NUM As Long

' 同じ形を選択する
Public IS_SELECT_FIGURE As Boolean

' 同じ大きさを選択する
Public IS_SELECT_SIZE As Boolean

' 同じ色を選択する
Public IS_SELECT_COLOR AS Boolean

' 図形の複製時のmargin
Public Const COLONE_MARGINE As Long = 2

' 初期化
Public Sub initModel()
  ' リボン起動時に初期化

  SELECT_SORT_ORDER = SORT_ORDER_ROW
  COLLISION_NUM_STRING = "0.50倍 (標準)"
  COLLISION = 0.50
  START_NUM = 1
  IS_CHANGE_TEXT   = TRUE
  IS_SELECT_FIGURE = TRUE
  IS_SELECT_COLOR  = TRUE
  IS_SELECT_SIZE   = TRUE
  ENABLED_SART_NUMBER = FALSE
End Sub
