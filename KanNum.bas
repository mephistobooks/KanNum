Attribute VB_Name = "KanNum"
'''' KanNum.bas
'
'
'Author: mephistobooks (http://voidptr.seesaa.net)
'Date: 2010 Aug. 13
'Updated: Sep. 1st, 2011 bugfix (��, �S, �\)
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
'  v �́A�������ŏ����ꂽ�A�Z���܂��͕�����
'
'
'DESCRIPTION
'  �������𐔒l�ɕϊ�����B���l���������ɕϊ�����֐��FNUMBERSTRING �̋t���s���B
'  �������́A���̒P�ʂ܂Ŏw��ł���B
'  �������̎w����@�́A�Z���܂��͕�����Ŏw��ł���B
'  �������̕\�L�́A���L�̗�̂�����ɂ��Ή��F
'
'  ��.
'       STRINGNUMBER("��㎵�Z") => 1,976
'   STRINGNUMBER("���S���\�Z") => 1,976
'        STRINGNUMBER("56��3��") => 563,000
'  STRINGNUMBER("�Q蔌ޕS�݈�E") => 35,000,010
'
'REFERENCES
'  Q.�gExcel�Ŋ������𐔒l�ɕϊ�������@�������Ă��������B�h
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
    '  assume that kanji-number string (������������) consists of three type of
    '  number string: compound units, semi-compound units, and literal
    '  numbers.
    '
    '  compound units (��,��,��,��) consists of semi-compound units
    '  and literal number string.
    '
    '  semi-compound units (��,�S,�\) consists of themselves (ex. ��\),
    '  and also they sometimes includes literal number string (ex. 3��4�S).
    '
    '  literal number string (�Z�`��) is able to convert to number (value)
    '  directly.
    '
    '*Way of getting the value of kanji-number string
    '  generate the expression from kanji-number string like the following:
    '    ((n_1)��)+(n_2)��+(n_3))
    '  where n_? is a*1000+b*100+c*10+d, and a-d here indicates 0-9.
    '
    '  after the step (4), we can obtain the above expression which
    '  consists from only numbers, parentheses, product(*), and addition(+).
    '
    '  and finally, the value is obtained by evaluation of the generated
    '  expression at the step (5)
    '
    
    '''' (1) Normalize the value.
    '�@�@`normalize' means that hankaku, and all characters of
    ' [�O-�X],[�Z-��] to [0-9].
    '
    str = StrConv(str, vbNarrow) '���p
    
    'for semi-compound units
    str = Replace(str, "�E", "�\")
    str = Replace(str, "�", "��")
    str = Replace(str, "��", "��")
        
    'for the literal number strings.
    str = Replace(str, "��", "9")
    str = Replace(str, "��", "8")
    str = Replace(str, "��", "7")
    str = Replace(str, "�Z", "6")
    
    str = Replace(str, "��", "5")
    str = Replace(str, "��", "5")
    
    str = Replace(str, "�l", "4")
    
    str = Replace(str, "�O", "3")
    str = Replace(str, "�Q", "3")
    
    str = Replace(str, "��", "2")
    str = Replace(str, "��", "2")
    
    str = Replace(str, "��", "1")
    str = Replace(str, "��", "1")
    
    str = Replace(str, "�Z", "0")
    
    
    '''' (2) Structurize for the compound units: ��,��,��,��.
    '
    '  (.*)��+(.*)��+(.*)��+(.*)
    '
    str = "(" + str
    
    '
    str = Replace(str, "��", ")��+(")
    str = Replace(str, "��", ")��+(")
    str = Replace(str, "��", ")��+(")
    str = Replace(str, "��", ")��+(")
      
    str = str + ")"
    
       
    ''' (3) Numerize the value
    
    'semi-compound units
    str = Replace(str, "��", "*1000+")
    str = Replace(str, "�S", "* 100+")
    str = Replace(str, "�\", "*  10+")
    
    'compound units
    str = Replace(str, "��", "*10000000000000000+")
    str = Replace(str, "��", "*    1000000000000+")
    str = Replace(str, "��", "*        100000000+")
    str = Replace(str, "��", "*            10000+")
    
    
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

