object frmPrincipal: TfrmPrincipal
  Left = 0
  Top = 0
  Caption = 'frmPrincipal'
  ClientHeight = 320
  ClientWidth = 387
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object TrayIcon: TTrayIcon
    PopupMenu = PopMenu
    Visible = True
    Left = 16
    Top = 16
  end
  object PopMenu: TPopupMenu
    Left = 64
    Top = 16
    object Backup1: TMenuItem
      Caption = 'Backup'
      OnClick = Backup1Click
    end
    object Restaurao1: TMenuItem
      Caption = 'Restaura'#231#227'o'
      OnClick = Restaurao1Click
    end
    object Reparo1: TMenuItem
      Caption = 'Reparo'
      OnClick = Reparo1Click
    end
    object Fechar1: TMenuItem
      Caption = 'Fechar'
      OnClick = Fechar1Click
    end
  end
end
