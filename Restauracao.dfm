object frmRestauracao: TfrmRestauracao
  Left = 0
  Top = 0
  BorderStyle = bsSingle
  Caption = 'Restaura'#231#227'o'
  ClientHeight = 449
  ClientWidth = 595
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
  object MemoLog: TMemo
    Left = 8
    Top = 79
    Width = 577
    Height = 298
    Lines.Strings = (
      'MemoLog')
    TabOrder = 0
  end
  object ProgressBar: TProgressBar
    Left = 8
    Top = 383
    Width = 577
    Height = 25
    TabOrder = 1
  end
  object grp1: TGroupBox
    Left = 8
    Top = 8
    Width = 577
    Height = 65
    Caption = 'Par'#226'metros'
    TabOrder = 2
    object lbl1: TLabel
      Left = 328
      Top = 16
      Width = 150
      Height = 13
      Caption = 'Selecione o Arquivo de Backup:'
    end
    object txtCaminhoArq: TEdit
      Left = 328
      Top = 32
      Width = 209
      Height = 21
      TabOrder = 0
    end
    object btProcura: TBitBtn
      Left = 536
      Top = 30
      Width = 25
      Height = 25
      Caption = '...'
      DoubleBuffered = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentDoubleBuffered = False
      ParentFont = False
      TabOrder = 1
      OnClick = btProcuraClick
    end
  end
  object btnInicia: TBitBtn
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
    TabOrder = 3
    OnClick = btnIniciaClick
  end
  object btnSair: TBitBtn
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
    TabOrder = 4
    OnClick = btnSairClick
  end
  object Thread: TJvThread
    Exclusive = True
    MaxCount = 0
    RunOnCreate = True
    FreeOnTerminate = True
    OnExecute = ThreadExecute
    Left = 16
    Top = 96
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
    Left = 136
    Top = 96
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Arquivo Zipado (*.zip)|*.zip|Todos os Arquivos (*.*)|*.*'
    Left = 192
    Top = 104
  end
  object IBRestoreService1: TIBRestoreService
    TraceFlags = []
    PageBuffers = 0
    Left = 80
    Top = 120
  end
  object IBBackupService: TIBBackupService
    TraceFlags = []
    BlockingFactor = 0
    Options = []
    Left = 256
    Top = 112
  end
end
