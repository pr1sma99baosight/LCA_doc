object aboutform: Taboutform
  Left = 739
  Top = 321
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'About'
  ClientHeight = 332
  ClientWidth = 389
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 32
    Top = 8
    Width = 337
    Height = 273
    BevelInner = bvRaised
    BevelOuter = bvLowered
    Color = clWhite
    TabOrder = 0
    object ProductName: TLabel
      Left = 64
      Top = 40
      Width = 193
      Height = 21
      Caption = #23453#38050#20135#21697#29983#21629#21608#26399#35780#20215#36719#20214
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #24494#36719#38597#40657
      Font.Style = []
      ParentFont = False
      IsControl = True
    end
    object Version: TLabel
      Left = 112
      Top = 88
      Width = 70
      Height = 21
      Caption = #29256#26412#65306'4.0'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #24494#36719#38597#40657
      Font.Style = []
      ParentFont = False
      IsControl = True
    end
    object Copyright: TLabel
      Left = 48
      Top = 232
      Width = 239
      Height = 21
      Alignment = taCenter
      Caption = #23453#23665#38050#38081#32929#20221#26377#38480#20844#21496'   '#29256#26435#25152#26377
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #24494#36719#38597#40657
      Font.Style = []
      ParentFont = False
      IsControl = True
    end
  end
  object Button1: TButton
    Left = 136
    Top = 288
    Width = 129
    Height = 33
    Caption = 'OK'
    TabOrder = 1
    OnClick = Button1Click
  end
end
