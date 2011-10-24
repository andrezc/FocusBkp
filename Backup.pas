unit Backup;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Principal, StdCtrls, Buttons, JvComponentBase, JvThread, IBServices,
  ComCtrls, DUtilit, ZipForge;

type
  TfrmBackup = class(TForm)
    grp1: TGroupBox;
    MemoLog: TMemo;
    txtCaminho: TEdit;
    btSeleciona: TBitBtn;
    lbl1: TLabel;
    cbxCBRecoLixo: TCheckBox;
    cbxCBTran: TCheckBox;
    cbxCBIgnoChec: TCheckBox;
    cbxCBIgnoLimb: TCheckBox;
    cbxCBDetalhes: TCheckBox;
    btInicia: TBitBtn;
    btSair: TBitBtn;
    OpenDialog1: TOpenDialog;
    ThreadBackup: TJvThread;
    IBBackupService1: TIBBackupService;
    ProgressBar1: TProgressBar;
    ZipForge1: TZipForge;
    procedure FormCreate(Sender: TObject);
    procedure btSairClick(Sender: TObject);
    procedure btSelecionaClick(Sender: TObject);
    procedure btIniciaClick(Sender: TObject);
    procedure ThreadBackupExecute(Sender: TObject; Params: Pointer);
  private
    procedure GeraBackup;
    function StrToBool(const Valor: string;
  const ValoVerd: String='Sim'): Boolean;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmBackup: TfrmBackup;
  archiver : TZipForge;

implementation

{$R *.dfm}

function TfrmBackup.StrToBool(const Valor: string;
  const ValoVerd: String='Sim'): Boolean;
begin
  Result:=Valor=ValoVerd;
end;

procedure TfrmBackup.GeraBackup;
begin
  MemoLog.Clear;
  if txtCaminho.Text = Trim('') then
  begin
    MessageBox(0, 'Selecione um Banco de Dados.', 'Atenção!', MB_ICONWARNING or MB_OK);
    Exit;
  end;

  MemoLog.Lines.Add('Preparando para gerar o backup...');
  with IBBackupService1 do
  begin
    DatabaseName:= txtCaminho.Text;
    ServerName:= 'localhost';
    BackupFile.Clear;
    BackupFile.Add(DiretorioDoExecutavel+'FOCUS_BACKUP\Backup_Focus_.fbk');
    Params.Clear;
    Params.Add('user_name=SYSDBA');
    Params.Add('password=masterkey');
    Options:=[];
    if cbxCBIgnoChec.Checked then
      Options:=Options+[IgnoreChecksums];
    if cbxCBIgnoLimb.Checked then
      Options:=Options+[IgnoreLimbo];
    if not cbxCBRecoLixo.Checked then
      Options:=Options+[NoGarbageCollection];
    if not cbxCBTran.Checked then
      Options:=Options+[NonTransportable];
    Verbose:= cbxCBDetalhes.Checked;

    MemoLog.Lines.Add('-  Ignorar checksum: '+BoolToStr(cbxCBIgnoChec.Checked));
    MemoLog.Lines.Add('-  Ignorar transações em limbo: '+BoolToStr(cbxCBIgnoLimb.Checked));
    MemoLog.Lines.Add('-  Coletar lixo: '+BoolToStr(cbxCBRecoLixo.Checked));
    MemoLog.Lines.Add('-  Formato transportável: '+BoolToStr(cbxCBTran.Checked));
    MemoLog.Lines.Add('-  Verbose: '+BoolToStr(cbxCBTran.Checked));
    MemoLog.Lines.Add('-  Nome do servidor: '+ServerName);
    MemoLog.Lines.Add('');
    Active:=True;
    MemoLog.Lines.Add('');
    MemoLog.Lines.Add('                     Inicio                    ');
    MemoLog.Lines.Add('--------------------------------------------');
    Application.ProcessMessages;
    MemoLog.Lines.Add('');

    try
      btInicia.Enabled:= False;
      Screen.Cursor:= crHourGlass;
      ServiceStart;
      while not EoF do
      begin
        MemoLog.Lines.Add(GetNextLine);
        ProgressBar1.Position:= ProgressBar1.Position + 1;
      end;
      MemoLog.Lines.Add(' ');
      MemoLog.Lines.Add('Iniciando a compactação...');
      MemoLog.Lines.Add(' ');

      archiver := TZipForge.Create(nil);

      with archiver do
      begin
        FileName := DiretorioDoExecutavel+'FOCUS_BACKUP\Focus_bkp_'+FormatDateTime('yyyy-mm-dd-hhnn',now)+'.zip';
        OpenArchive(fmCreate);
        BaseDir :=DiretorioDoExecutavel+'FOCUS_BACKUP\';
        AddFiles(DiretorioDoExecutavel+'FOCUS_BACKUP\Backup_Focus_.fbk');
        CloseArchive();
        deletefile(DiretorioDoExecutavel+'FOCUS_BACKUP\Backup_Focus_.fbk');
        while not Eof do
        begin
          MemoLog.Lines.Add(GetNextLine);
        end;

      end

    finally
      Active:= False;
      MemoLog.Lines.Add('');
      MemoLog.Lines.Add('--------------------------------------------');
      MemoLog.Lines.Add('                     Fim                    ');
      btInicia.Enabled:= True;
      Screen.Cursor:= crDefault;
    end;

    MemoLog.Lines.Add('');
    MemoLog.Lines.Add('');
    MemoLog.Lines.Add('Backup Concluído.');
    MessageBox(0, 'Backup Concluído!    ', 'Sucesso!', MB_ICONINFORMATION or MB_OK);
  end;
end;

procedure TfrmBackup.ThreadBackupExecute(Sender: TObject; Params: Pointer);
begin
  GeraBackup;
end;

procedure TfrmBackup.btIniciaClick(Sender: TObject);
begin
  ThreadBackup.Priority:= tpLower;
  ThreadBackup.ExecuteAndWait(Self);
end;

procedure TfrmBackup.btSairClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmBackup.btSelecionaClick(Sender: TObject);
begin
  OpenDialog1.Execute();
  txtCaminho.Text:= OpenDialog1.FileName;
end;

procedure TfrmBackup.FormCreate(Sender: TObject);
begin
  MemoLog.Clear;
  ProgressBar1.Position:= 0;
end;


end.
