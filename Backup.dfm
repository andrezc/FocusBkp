object frmBackup: TfrmBackup
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Backup'
  ClientHeight = 447
  ClientWidth = 593
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object grp1: TGroupBox
    Left = 8
    Top = 8
    Width = 577
    Height = 65
    Caption = 'Par'#226'metros'
    TabOrder = 0
    object lbl1: TLabel
      Left = 377
      Top = 14
      Width = 140
      Height = 13
      Caption = 'Caminho do Banco de Dados:'
    end
    object txtCaminho: TEdit
      Left = 377
      Top = 33
      Width = 177
      Height = 21
      TabOrder = 0
    end
    object btSeleciona: TBitBtn
      Left = 549
      Top = 31
      Width = 25
      Height = 23
      Caption = '...'
      DoubleBuffered = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentDoubleBuffered = False
      ParentFont = False
      TabOrder = 1
      OnClick = btSelecionaClick
    end
    object cbxCBRecoLixo: TCheckBox
      Left = 8
      Top = 19
      Width = 81
      Height = 17
      Caption = 'Recolher lixo'
      Checked = True
      State = cbChecked
      TabOrder = 2
    end
    object cbxCBTran: TCheckBox
      Left = 8
      Top = 42
      Width = 89
      Height = 17
      Caption = 'Transport'#225'vel'
      Checked = True
      State = cbChecked
      TabOrder = 3
    end
    object cbxCBIgnoChec: TCheckBox
      Left = 103
      Top = 19
      Width = 145
      Height = 17
      Caption = 'Ignorar erros de checksum'
      Enabled = False
      TabOrder = 4
    end
    object cbxCBIgnoLimb: TCheckBox
      Left = 210
      Top = 45
      Width = 153
      Height = 17
      Caption = 'Ignorar transa'#231#245'es em limbo'
      TabOrder = 5
    end
    object cbxCBDetalhes: TCheckBox
      Left = 103
      Top = 45
      Width = 89
      Height = 17
      Caption = 'Detalhamento'
      Checked = True
      State = cbChecked
      TabOrder = 6
    end
  end
  object MemoLog: TMemo
    Left = 8
    Top = 79
    Width = 577
    Height = 298
    Lines.Strings = (
      'MemoLog')
    TabOrder = 1
  end
  object btInicia: TBitBtn
    Left = 248
    Top = 415
    Width = 179
    Height = 25
    Caption = 'INICIAR'
    DoubleBuffered = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentDoubleBuffered = False
    ParentFont = False
    TabOrder = 2
    OnClick = btIniciaClick
  end
  object btSair: TBitBtn
    Left = 433
    Top = 415
    Width = 152
    Height = 25
    Caption = 'SAIR'
    DoubleBuffered = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentDoubleBuffered = False
    ParentFont = False
    TabOrder = 3
    OnClick = btSairClick
  end
  object ProgressBar1: TProgressBar
    Left = 8
    Top = 383
    Width = 577
    Height = 25
    TabOrder = 4
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Banco de Dados Focus (*.FOC)|*.foc|Todos os arquivos (*.*)|*.*'
    Left = 128
    Top = 112
  end
  object ThreadBackup: TJvThread
    Exclusive = True
    MaxCount = 0
    RunOnCreate = True
    FreeOnTerminate = True
    OnExecute = ThreadBackupExecute
    Left = 176
    Top = 112
  end
  object IBBackupService1: TIBBackupService
    ServerName = 'localhost'
    Protocol = TCP
    LoginPrompt = False
    TraceFlags = []
    BlockingFactor = 0
    Options = []
    Left = 216
    Top = 112
  end
  object ZipForge1: TZipForge
    ExtractCorruptedFiles = False
    CompressionLevel = clFastest
    CompressionMode = 1
    CurrentVersion = '5.03 '
    SpanningMode = smNone
    SpanningOptions.AdvancedNaming = False
    SpanningOptions.VolumeSize = vsAutoDetect
    Options.FlushBuffers = True
    Options.OEMFileNames = True
    InMemory = False
    Zip64Mode = zmDisabled
    UnicodeFilenames = False
    EncryptionMethod = caPkzipClassic
    Left = 280
    Top = 112
  end
end
