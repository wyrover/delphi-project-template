object Demo3Form: TDemo3Form
  Left = 0
  Top = 0
  Caption = 'Demo3Form'
  ClientHeight = 591
  ClientWidth = 751
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object spl1: TSplitter
    Left = 0
    Top = 177
    Width = 751
    Height = 3
    Cursor = crVSplit
    Align = alTop
    ExplicitWidth = 123
  end
  object vTree: TVirtualStringTree
    Left = 0
    Top = 0
    Width = 751
    Height = 177
    Align = alTop
    Header.AutoSizeIndex = 0
    Header.Font.Charset = DEFAULT_CHARSET
    Header.Font.Color = clWindowText
    Header.Font.Height = -11
    Header.Font.Name = 'Tahoma'
    Header.Font.Style = []
    Header.MainColumn = -1
    TabOrder = 0
    OnFreeNode = vTreeFreeNode
    OnGetText = vTreeGetText
    Columns = <>
  end
  object btn1: TButton
    Left = 88
    Top = 224
    Width = 75
    Height = 25
    Caption = 'btn1'
    TabOrder = 1
    OnClick = btn1Click
  end
  object btn2: TButton
    Left = 200
    Top = 224
    Width = 75
    Height = 25
    Caption = 'btn2'
    TabOrder = 2
    OnClick = btn2Click
  end
  object btn3: TButton
    Left = 312
    Top = 224
    Width = 75
    Height = 25
    Caption = 'btn3'
    TabOrder = 3
    OnClick = btn3Click
  end
  object btn4: TButton
    Left = 448
    Top = 224
    Width = 75
    Height = 25
    Caption = 'Test AutoIE'
    TabOrder = 4
    OnClick = btn4Click
  end
  object EmbeddedWB1: TEmbeddedWB
    Left = 88
    Top = 320
    Width = 417
    Height = 201
    TabOrder = 5
    Silent = False
    DisableCtrlShortcuts = 'N'
    UserInterfaceOptions = [EnablesFormsAutoComplete, EnableThemes]
    About = ' EmbeddedWB http://bsalsa.com/'
    PrintOptions.HTMLHeader.Strings = (
      '<HTML></HTML>')
    PrintOptions.Orientation = poPortrait
    ControlData = {
      4C000000192B0000C61400000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126208000000000000004C0000000114020000000000C000000000000046
      8000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000}
  end
end
