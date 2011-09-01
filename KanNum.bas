Attribute VB_Name = "KanNum"
'''' KanNum.bas
'
'
'Author: mephistobooks (http://voidptr.seesaa.net)
'Date: 2010 Aug. 13
'Updated: Sep. 1st, 2011 bugfix (千, 百, 十)
'


''' as use stricts in Perl;
Option Explicit

'Private Const DEBUG_KANNUM = True
Private Const DEBUG_KANNUM = False


'NAME
'  STRINGNUMBER
'
'SYNOPSIS
'  =STRINGNUMBER(v)
'  v は、漢数字で書かれた、セルまたは文字列
'
'
'DESCRIPTION
'  漢数字を数値に変換する。数値を漢数字に変換する関数：NUMBERSTRING の逆を行う。
'  漢数字は、京の単位まで指定できる。
'  漢数字の指定方法は、セルまたは文字列で指定できる。
'  漢数字の表記は、下記の例のいずれにも対応：
'
'  例.
'       STRINGNUMBER("一九七六") => 1,976
'   STRINGNUMBER("千九百七十六") => 1,976
'        STRINGNUMBER("56万3千") => 563,000
'  STRINGNUMBER("参阡伍百萬壱拾") => 35,000,010
'
'REFERENCES
'  Q.“Excelで漢数字を数値に変換する方法を教えてください。”
'  http://q.hatena.ne.jp/1268555767
'
Public Function STRINGNUMBER(ByVal varcl As Variant) As Variant
    '''
    Dim str As String       'kanji-number string (working variable)
    Dim str_org As String   'kanji-number string (original)
    
    '
    Dim tmp As String
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''' (0) Initialize the variables
    '
    tmp = TypeName(varcl)
    
    If DEBUG_KANNUM Then
        Call MsgBox("TypeName(varcl)[" & tmp & "]" & vbCrLf & _
                    "str[" & str & "]")
    End If
    
    
    ' Process the argument due to its type.
    If TypeName(tmp) = "String" Then
        str_org = varcl
    Else
        str_org = varcl.Value
    End If
    
    'Kanji-number string.
    str = str_org
        
    If DEBUG_KANNUM Then
        Call MsgBox("TypeName(varcl)[" & tmp & "]" & vbCrLf & _
                    "str[" & str & "]")
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Algorithm:
    '
    '*Assumption
    '  assume that kanji-number string (漢数字文字列) consists of three type of
    '  number string: compound units, semi-compound units, and literal
    '  numbers.
    '
    '  compound units (京,兆,億,万) consists of semi-compound units
    '  and literal number string.
    '
    '  semi-compound units (千,百,十) consists of themselves (ex. 千十),
    '  and also they sometimes includes literal number string (ex. 3千4百).
    '
    '  literal number string (〇～九) is able to convert to number (value)
    '  directly.
    '
    '*Way of getting the value of kanji-number string
    '  generate the expression from kanji-number string like the following:
    '    ((n_1)億)+(n_2)万+(n_3))
    '  where n_? is a*1000+b*100+c*10+d, and a-d here indicates 0-9.
    '
    '  after the step (4), we can obtain the above expression which
    '  consists from only numbers, parentheses, product(*), and addition(+).
    '
    '  and finally, the value is obtained by evaluation of the generated
    '  expression at the step (5)
    '
    
    '''' (1) Normalize the value.
    '　　`normalize' means that hankaku, and all characters of
    ' [０-９],[〇-九] to [0-9].
    '
    str = StrConv(str, vbNarrow) '半角
    
    'for semi-compound units
    str = Replace(str, "拾", "十")
    str = Replace(str, "阡", "千")
    str = Replace(str, "萬", "万")
        
    'for the literal number strings.
    str = Replace(str, "九", "9")
    str = Replace(str, "八", "8")
    str = Replace(str, "七", "7")
    str = Replace(str, "六", "6")
    
    str = Replace(str, "五", "5")
    str = Replace(str, "伍", "5")
    
    str = Replace(str, "四", "4")
    
    str = Replace(str, "三", "3")
    str = Replace(str, "参", "3")
    
    str = Replace(str, "二", "2")
    str = Replace(str, "弐", "2")
    
    str = Replace(str, "一", "1")
    str = Replace(str, "壱", "1")
    
    str = Replace(str, "〇", "0")
    
    
    '''' (2) Structurize for the compound units: 京,兆,億,万.
    '
    '  (.*)兆+(.*)億+(.*)万+(.*)
    '
    str = "(" + str
    
    '
    str = Replace(str, "京", ")京+(")
    str = Replace(str, "兆", ")兆+(")
    str = Replace(str, "億", ")億+(")
    str = Replace(str, "万", ")万+(")
      
    str = str + ")"
    
       
    ''' (3) Numerize the value
    
    'semi-compound units
    str = Replace(str, "千", "*1000+")
    str = Replace(str, "百", "* 100+")
    str = Replace(str, "十", "*  10+")
    
    'compound units
    str = Replace(str, "京", "*10000000000000000+")
    str = Replace(str, "兆", "*    1000000000000+")
    str = Replace(str, "億", "*        100000000+")
    str = Replace(str, "万", "*            10000+")
    
    
    ''' (4) correct the expression of the value.
    str = Replace(str, "()", "0")
    
    str = Replace(str, "++", "+")
    str = Replace(str, "+)", "+0)")
          
    str = Replace(str, "(*", "(1*")
    str = Replace(str, "+*", "+1*")
    
    If DEBUG_KANNUM Then
        Call MsgBox("org:" & str_org & vbCrLf & "str:" & str)
    End If
    
    
    ''' (5) Eval the generated expression!
    STRINGNUMBER = Application.Evaluate(str)
    
End Function

