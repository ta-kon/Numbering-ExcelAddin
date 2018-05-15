連番太郎 (Excel add-in)
====

[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](https://raw.githubusercontent.com/ta-kon/Numbering-ExcelAddin/master/LICENSE)

選択した図形を採番するExcelアドイン

## 目次

- [1. Description](#1-description)
- [2. Demo](#2-demo)
    - [2.1. 採番と自動選択](#21-採番と自動選択)
    - [2.2. 図形の複製](#22-図形の複製)
- [3. VS.](#3-vs)
- [4. Requirement](#4-requirement)
- [5. Usage](#5-usage)
    - [5.1. 採番](#51-採番)
    - [5.2. 選択](#52-選択)
    - [5.3. 複製](#53-複製)
    - [5.4. テキストの除去](#54-テキストの除去)
- [6. Install](#6-install)
- [7. Uninstall](#7-uninstall)
- [8. Thanks](#8-thanks)
- [9. Licence](#9-licence)

## 1. Description
Excelで以下の作業が簡単になります。
* 図形の採番
* 図形のテキストの除去
* 図形の複製
* 図形の選択

本ソフトをインストールすることで、  
Excelに「連番太郎」というリボンが作成されます。  
そこから使いたい機能をクリックするだけで実行できます。

## 2. Demo
アニメGIFを貼付けて実際の動作例を見せます。

### 2.1. 採番と自動選択
![NumberingShape](https://raw.githubusercontent.com/ta-kon/Numbering-ExcelAddin/master/movie/NumberingShape.gif)

### 2.2. 図形の複製
![CloneShape](https://raw.githubusercontent.com/ta-kon/Numbering-ExcelAddin/master/movie/CloneShape.gif)

## 3. VS. 
手動で一つひとつ選択して採番していた作業が
本アドインが代わりに行ってくれます。

選択した順に採番することも可能です。

採番順序の際に *挿入ソート* を採用しております。

## 4. Requirement
以下の環境で動作確認をしております。  
* Excel 2010
* Excel 2016

※ Windows版Excelのみ動作確認済み

## 5. Usage
使い方

### 5.1. 採番
1. 採番したい図形を複数クリック
2. 採番開始をクリック

### 5.2. 選択
1. 複数選択したい図形を*1つだけ*クリック
2. 自動選択開始をクリック

### 5.3. 複製
1. 複製したい図形を*1つだけ*クリック
2. 図形の複製をクリック

### 5.4. テキストの除去
1. テキストを除去したい図形を複数クリック
2. テキストの除去をクリック

## 6. Install

1. 以下の場所から「Source code (zip)」をダウンロード  
https://github.com/ta-kon/Numbering-ExcelAddin/releases

2. ダウンロードしたzipファイルを展開し、
展開したフォルダ内にある  
Install.vbs を実行

3. インストール完了  
Excelのリボンに「連番太郎」が自動で追加されます。

※ インストール後は、展開したファイルは削除しても大丈夫です。

## 7. Uninstall
アンインストール方法

Uninstall.vbs を実行

## 8. Thanks

以下のサイトを参考にしました。  

[Excel VBAコーディング ガイドライン案 - Qiita](https://qiita.com/mima_ita/items/8b0eec3b5a81f168822d)  
VBAコーディングの目安にさせて頂きました。  

[Try #008 – VBAのモダンな開発環境を構築してみた | dayjournal memo](https://day-journal.com/memo/try-008/)  
VBAをVS Codeでコーディングする際に参考にしました。

[tcsh/text-scripting-vba: Modules for text scripting on VBA](https://github.com/tcsh/text-scripting-vba)  
VBAをVS Codeでコーディングする際に使用したマクロ。これが無かったら、開発に挫折していました。  

[IRibbonUIオブジェクトがNothingになったときの対処法](http://www.ka-net.org/ribbon/ri64.html)  
リボンUIを利用する際に、オブジェクトがNothingになったので参考にしました。  

[ある SE のつぶやき - VBScript で Excel にアドインを自動でインストール/アンインストールする方法](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)  
Install.vbs / Uninstall.vbs に使用しております。  

## 9. Licence
本プログラムはフリーソフトウェアです。  
個人・法人に限らず利用者は自由に使用および配布することができます。  

本プログラムは無償で利用できますが、  
作者は本プログラムの使用にあたり生じる障害や問題に対して一切の責任を負いません。  

ソースを利用する場合にはMITライセンスです。
