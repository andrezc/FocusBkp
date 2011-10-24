unit Restauracao;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, JvComponentBase, JvThread, Buttons, IBServices,
  DUtilit, ZipForge;

type
  TfrmRestauracao = class(TForm)
    MemoLog: TMemo;
    ProgressBar: TProgressBar;
    Thread: TJvThread;
    grp1: TGroupBox;
    txtCaminhoArq: TEdit;
    btProcura: TBitBtn;
    lbl1: TLabel;
    btnInicia: TBitBtn;
    btnSair: TBitBtn;
    ZipForge1: TZipForge;
    OpenDialog1: TOpenDialog;
    IBRestoreService1: TIBRestoreService;
    IBBackupService: TIBBackupService;
    procedure FormCreate(Sender: TObject);
    procedure btnSairClick(Sender: TObject);
    procedure btnIniciaClick(Sender: TObject);
    procedure btProcuraClick(Sender: TObject);
    procedure ThreadExecute(Sender: TObject; Params: Pointer);
  private
    procedure Restaurabkp;
    procedure Informa(mensagem: string);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRestauracao: TfrmRestauracao;
  archiver : TZipForge;

implementation

{$R *.dfm}

procedure TfrmRestauracao.Informa (mensagem: string);
begin
  MemoLog.Lines.Add(mensagem);
  Application.ProcessMessages;
end;

function DirBkp : String;
begin
  Result:= DiretorioDoExecutavel+'FOCUS_BACKUP\';
end;

function GetWindowsDrive: Char;
var
  S: string;
begin
  SetLength(S, MAX_PATH);
  if GetWindowsDirectory(PChar(S), MAX_PATH) > 0 then
    Result := string(S)[1]
  else
    Result := #0;
end;

procedure TfrmRestauracao.Restaurabkp;
var
  Caminho,
  DirCort,
  DirCort2 : string;
begin
  if txtCaminhoArq.Text = Trim('') then
  begin
    Informa('********** Aten��o! **********');
    Informa('Voc� n�o informou o caminho do Arquivo de Backup.');
    Exit;
  end;

  MemoLog.Clear;
  btnInicia.Enabled:=False;
  MemoLog.Lines.Add('Preparando para iniciar a Restaura��o...');


  if(FileExists(DirBkp + 'FOCUS.FOC')) then
  begin
    RenameFile(dirBkp + 'FOCUS.FOC', dirBkp + 'FOCUS_OLD.FOC');

    try
      with IBBackupService do
      begin
        DatabaseName:= dirBkp + 'FOCUS.FOC';
        ServerName  :='localhost';
        BackupFile.Clear;
        BackupFile.Add(dirBkp + 'FOCUS.FOC');
        Protocol    :=TCP;
        Params.Clear;
        Params.Add('user_name=SYSDBA');
        Params.Add('password=masterkey');
        LoginPrompt := False;
        Active      := True;

        Application.ProcessMessages;

        try
          ServiceStart;
          while not Eof do
          begin
            Informa(GetNextLine);
          end;
        except
          on e : Exception do
            Informa(e.Message);
        end;
        Active:=False;
        Informa('');
        Informa('********* Restaura��o Conclu�da! *********');
      end;

        //DMBanco.IBDB.Connected:=True;//conecta o sistema na base de dados
        btnInicia.Enabled:=True;
    except
      on E: Exception do
      begin
        //  DMBanco.IBDB.Connected:=True;//conecta o sistema na base de dados
        btnInicia.Enabled:=True;
      end;
    end;

    // Compacta��o
    archiver := TZipForge.Create(nil);
    try
      with archiver do
      begin
        FileName :=dirBkp + 'FOCUS_OLD.zip';
        OpenArchive(fmCreate);
        BaseDir :=dirBkp;
        caminho := DirBkp + 'FOCUS_OLD.FOC';
        AddFiles(caminho);
        CloseArchive();
        deletefile(caminho);
      end
    except
      on e : Exception do
      begin
        Informa('***** Ocorreu algum problema ao compactar. *****');
        Informa(e.Message);
        Exit;
      end;
    end;
  end;

  if (UpperCase(TiraAPartirDoUltimo(OpenDialog1.FileName,'.')))='ZIP' then
  begin
    archiver := TZipForge.Create(nil);
    try
      with archiver do
      begin
        CopiaArquivos(txtCaminhoArq.Text, GetWindowsDrive+':\', '');
        FileName := txtCaminhoArq.Text;
        OpenArchive(fmOpenReadWrite);
        Archiver.RenameFile('*.fbk', 'DADOS.fbk');
        BaseDir := DirBkp ;
        ExtractFiles('*.*');
        CloseArchive();
      end;

    except
      on e : Exception do
      begin
        Informa('***** Ocorreu um erro ao descompactar o Arquivo de Backup. *****');
        Informa(e.Message);
        Exit;
      end;
    end ;
  end;

  try
    with IBRestoreService1 do
    begin
      Informa('');
      Informa('***** Procurando o arquivo "DADOS.fbk"... ******');
      Sleep(6000);
      if not FileExists (dirBkp + 'DADOS.fbk') then
      begin
        Informa('O arquivo n�o foi encontrado!');
        Informa('');
        btnInicia.Enabled:= False;
        dirCort := TiraAteOUltimo(txtCaminhoArq.Text,'\');
        dirCort2:= dirCort+'\DADOS.fbk';
        RenameFile(txtCaminhoArq.Text,dirCort2);
      end;

      ServerName:= 'localhost';
      loginPrompt:= False;
      Params.Add('user_name=SYSDBA');
      Params.Add('password=masterkey');
      Active := true;
      Verbose := true;
      DatabaseName.Add (DirBkp + 'FOCUS.FOC');
      BackupFile.Add(dirBkp + 'DADOS.fbk');
      Informa('*****************  Inicio *****************');
      Application.ProcessMessages;
      Informa('');
      with IBRestoreService1 do
      begin
        ServiceStart;
        while not Eof do
        begin
          Informa(GetNextLine);
        end;
        Active:= False;
        Informa('***************** Fim *****************');
        Options:= [];
      end;
    end;

  finally
    Screen.Cursor:= crDefault;
    txtCaminhoArq.Clear;
    btnInicia.Enabled := True;
  end;
end;


procedure TfrmRestauracao.ThreadExecute(Sender: TObject; Params: Pointer);
begin
  Restaurabkp;
end;

procedure TfrmRestauracao.btnIniciaClick(Sender: TObject);
begin
  Thread.Priority:= tpLower;
  Thread.ExecuteAndWait(Self);
end;

procedure TfrmRestauracao.btnSairClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmRestauracao.btProcuraClick(Sender: TObject);
begin
  OpenDialog1.Execute();
  txtCaminhoArq.Text:= OpenDialog1.FileName;
end;

procedure TfrmRestauracao.FormCreate(Sender: TObject);
begin
  MemoLog.Clear;
  ProgressBar.Position:= 0;
end;
end.
