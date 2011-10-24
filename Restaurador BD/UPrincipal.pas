unit UPrincipal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, ShellAPI, FileCtrl, Registry, DBFunctions,
  IB_Components, IB_Access, IBODataset, IniFiles, DUtilit, JvComponentBase,
  JvEnterTab, UComp, JvExControls, JvAnimatedImage, JvGIFCtrl, XPMan;

type
  TfrmPrincipal = class(TForm)
    pnlSuperior: TPanel;
    pnlInferior: TPanel;
    lblInforma��o: TLabel;
    lblCaminhoFB: TLabel;
    edtCaminhoFB: TEdit;
    lblCaminhoBD: TLabel;
    edtCaminhoBDFocus: TEdit;
    btnPathFB: TSpeedButton;
    btnPathBD: TSpeedButton;
    IBODatabase1: TIBODatabase;
    OpenDialog1: TOpenDialog;
    pnlBotaoSair: TPanel;
    btnSair: TBitBtn;
    lblMens: TLabel;
    pnlBtnExecutar: TPanel;
    btnExecutar: TBitBtn;
    JvEnterAsTab1: TJvEnterAsTab;
    JvGIFAnimator1: TJvGIFAnimator;
    XPManifest1: TXPManifest;
    procedure btnSairClick(Sender: TObject);
    procedure btnPathFBClick(Sender: TObject);
    procedure btnPathBDClick(Sender: TObject);
    procedure btnExecutarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    function ShellExecuteAndWait(cArquivoPrograma, cParametros: string;
      iOpcaoJanela: integer): DWORD;
    procedure SalvaIni;
    procedure LeIni;
    function PegaCaminhoFBPeloRegistro: string;
    procedure ControlFBSvr(bStart: Boolean);
    procedure MostraMensagem(txt: string = '');
    procedure CriaBat(CaminhoFB, PATHBD: string);
    procedure DesativaComps;
    procedure AtivaComps;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

{$R *.dfm}
const
  SELDIRHELP = 1000;

procedure TfrmPrincipal.DesativaComps;
begin
  DesligaComponente(frmPrincipal,btnSair);
  DesligaComponente(frmPrincipal,btnExecutar);
  DesligaComponente(frmPrincipal,edtCaminhoFB);
  DesligaComponente(frmPrincipal,edtCaminhoBDFocus);
  DesligaComponente(frmPrincipal,btnPathFB);
  DesligaComponente(frmPrincipal,btnPathBD);
  JvGIFAnimator1.Visible := True;
  JvGIFAnimator1.Animate := True;
end;

procedure TfrmPrincipal.AtivaComps;
begin
  LigaComponente(frmPrincipal,btnSair);
  LigaComponente(frmPrincipal,btnExecutar);
  LigaComponente(frmPrincipal,edtCaminhoFB);
  LigaComponente(frmPrincipal,edtCaminhoBDFocus);
  LigaComponente(frmPrincipal,btnPathFB);
  LigaComponente(frmPrincipal,btnPathBD);
  JvGIFAnimator1.Visible := False;
  JvGIFAnimator1.Animate := False;
end;

procedure TfrmPrincipal.btnExecutarClick(Sender: TObject);
var
  DirFocus: string;
begin
  MostraMensagem('Aguarde Validando Informa��es...');

  if not IBODatabase1.Connected then begin
    if not Vazio(edtCaminhoBDFocus.Text) then begin
      if FileExists(edtCaminhoBDFocus.Text) then begin
        IBODatabase1.Database := edtCaminhoBDFocus.Text;
        IBODatabase1.Connected := True;
      end;
    end
    else
    begin
      ShowMessage('N�o foi poss�vel estabelecer a conex�o com o banco de dados, verifique!');
      Exit;
    end;
  end;
  if Vazio(edtCaminhoFB.Text) or Vazio(edtCaminhoBDFocus.Text) then begin
    if Vazio(edtCaminhoFB.Text) then
      ShowMessage('Preencha o caminho da pasta BIN do Firebird para prosseguir!');
    if Vazio(edtCaminhoBDFocus.Text) then
      ShowMessage('Preencha o caminho do banco de dados FOCUS.FOC para prosseguir!');
    Exit;
  end;
  if not FileExists(edtCaminhoFB.Text + '\instsvc.exe') then begin
    ShowMessage(edtCaminhoFB.Text + '\instsvc.exe');
    ShowMessage('Verifique o caminho da pasta BIN do Firebird e se existe o arquivo INSTSVC.EXE dentro dela!'+pl+'Caso contr�rio verifique a instala��o do Firebird!');
    Exit;
  end;
  if not FileExists(edtCaminhoBDFocus.Text) then begin
    ShowMessage('Arquivo FOCUS.FOC inexistente, preencha o caminho corretamente!');
    Exit;
  end;
  if fDB.GetConexoesSimultaneas(IBODatabase1) > 1 then begin
    ShowMessage('Ainda existe(m) '+ IntToStr(fDB.GetConexoesSimultaneas(IBODatabase1)-1) +' encerre as conex�es e tente novamente!');
    Exit;
  end;
  MostraMensagem('Aguarde Salvando Informa��es...');
  Screen.Cursor := crHourGlass;
  DesativaComps;
  SalvaIni;

  DirFocus := ExtractFilePath(edtCaminhoBDFocus.Text);
  CriaBat(edtCaminhoFB.Text,edtCaminhoBDFocus.Text);

  MostraMensagem('Desabilitando Firebird...');
  ControlFBSvr(False);
  Sleep(5000);
  MostraMensagem('Habilitando Firebird...');
  ControlFBSvr(True);
  Sleep(5000);

  Application.ProcessMessages;
  MostraMensagem('Executando Reparador...');
  try
    if FileExists(DirFocus+'REPARA.BAT') then
      ShellExecuteAndWait(DirFocus+'REPARA.BAT','',SW_HIDE)
    else
    begin
      Screen.Cursor := crDefault;
      AtivaComps;
      ShowMessage('Falha ao reparar arquivo, coloque o exe principal na pasta focus!');
      Exit;
    end;

    Screen.Cursor := crDefault;
    if FileExists(edtCaminhoBDFocus.Text) then
      MostraMensagem('Reparador Executado com SUCESSO!')
    else
    begin
      MostraMensagem('Ocorreu um erro, entre em contato com o suporte!');
      RenameFile(DirFocus+'Focus.ANT',edtCaminhoBDFocus.Text);
    end;
    Application.ProcessMessages;
  except
    MostraMensagem('Ocorreu um erro, tente novamente!');
    Screen.Cursor := crDefault;
    Application.ProcessMessages;
  end;
  AtivaComps;
end;

procedure TfrmPrincipal.MostraMensagem(txt: string = '');
begin
  lblMens.Caption := txt;
  Application.ProcessMessages;
end;

procedure TfrmPrincipal.SalvaIni;
var
  ArqIni: TIniFile;
  NomeArquivo: string;
begin
  NomeArquivo := ChangeFileExt(Application.ExeName,'.ini');
  if FileExists(NomeArquivo) then
    DeleteFile(NomeArquivo);
  ArqIni := TIniFile.Create(NomeArquivo);
  try
    with ArqIni do begin
      WriteString ('Caminho','BIN_Firebird', edtCaminhoFB.Text);
      WriteString ('Caminho','BD',           edtCaminhoBDFocus.Text);
    end;
  finally
    ArqIni.Free;
  end;
end;

procedure TfrmPrincipal.LeIni;
var
  ArqIni: TIniFile;
  NomeArquivo: string;
begin
  NomeArquivo := ChangeFileExt(Application.ExeName,'.ini');
  if FileExists(NomeArquivo) then begin
    ArqIni := TIniFile.Create(NomeArquivo);
    try
      with ArqIni do begin
        edtCaminhoFB.Text       := ReadString('Caminho','BIN_Firebird','');
        edtCaminhoBDFocus.Text  := ReadString('Caminho','BD','');
      end;
    finally
      ArqIni.Free;
    end;
  end;
  if vazio(edtCaminhoBDFocus.Text) then begin
    NomeArquivo := DiretorioDoExecutavel + 'Focus.ini';
    if FileExists(NomeArquivo) then begin
      ArqIni := TIniFile.Create(NomeArquivo);
      try
        with ArqIni do
          edtCaminhoBDFocus.Text := ReadString('CONFIGURACOES','BancoDeDados','');
      finally
        ArqIni.Free;
      end;
    end;
  end;
end;

procedure TfrmPrincipal.btnPathBDClick(Sender: TObject);
var
  Dir: string;
begin
  OpenDialog1.Title := 'Caminho do FOCUS.FOC';
  if Vazio(edtCaminhoBDFocus.Text) then
    OpenDialog1.InitialDir := DiretorioDoExecutavel
  else
    OpenDialog1.InitialDir := ExtractFilePath(edtCaminhoBDFocus.Text);
  if OpenDialog1.Execute then
    edtCaminhoBDFocus.Text := OpenDialog1.FileName;

  if FileExists(edtCaminhoBDFocus.Text) then begin
    IBODatabase1.Connected := False;
    IBODatabase1.Database := edtCaminhoBDFocus.Text;
    IBODatabase1.Connected := True;
  end;
end;

procedure TfrmPrincipal.btnPathFBClick(Sender: TObject);
var
  Dir: string;
begin
  if Length(edtCaminhoFB.Text) <= 0 then
     Dir := ExtractFileDir(application.ExeName)
  else
     Dir := edtCaminhoFB.Text;

  if SelectDirectory(Dir, [sdAllowCreate, sdPerformCreate, sdPrompt],SELDIRHELP) then
    edtCaminhoFB.Text := Dir;
end;

procedure TfrmPrincipal.btnSairClick(Sender: TObject);
begin
  if (MessageBox(0, 'Deseja encerrar o REPARADOR?', 'Confirma', MB_ICONQUESTION or MB_YESNO) = idYes) then
    Application.Terminate;
end;

procedure TfrmPrincipal.Button1Click(Sender: TObject);
begin
  CriaBat(edtCaminhoFB.Text,edtCaminhoBDFocus.Text);
end;

procedure TfrmPrincipal.FormCreate(Sender: TObject);
begin
  MostraMensagem('Focus Reparador');
  LeIni;
  if Vazio(edtCaminhoFB.Text) then begin
    edtCaminhoFB.Text := PegaCaminhoFBPeloRegistro;
  end;
  if not Vazio(edtCaminhoBDFocus.Text) then begin
    if FileExists(edtCaminhoBDFocus.Text) then begin
      IBODatabase1.Database := edtCaminhoBDFocus.Text;
      IBODatabase1.Connected := True;
    end;
  end;
end;

procedure TfrmPrincipal.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F2 : btnExecutarClick(self);
    VK_ESCAPE : btnSairClick(Self);
  end;
end;

function TfrmPrincipal.PegaCaminhoFBPeloRegistro: string;
begin
  Result := '';
  with TRegistry.Create do begin
    RootKey := HKEY_LOCAL_MACHINE;
    if OpenKey('SOFTWARE\Firebird Project\Firebird Server\Instances', False) then
      Result := ReadString('DefaultInstance') + 'bin'
    else
    if OpenKey('SOFTWARE\Wow6432Node\Firebird Project\Firebird Server\Instances', False) then
      Result := ReadString('DefaultInstance') + 'bin';
    CloseKey;
    Free;
    inherited;
  end;
end;

procedure TfrmPrincipal.ControlFBSvr(bStart: Boolean);
var
  szBuff: String;
begin
  begin
    szBuff := edtCaminhoFB.text + 'bin\instsvc.exe';
    if FileExists(edtCaminhoFB.Text + 'bin\instsvc.exe') then
    case bStart of
      True: ShellExecute(0, nil, PChar(szBuff), '-s start', nil, SW_HIDE);
      False: ShellExecute(0, nil, PChar(szBuff), '-s stop', nil, SW_HIDE);
    end;
  end;
end;

{----------------------------------------------------------
* Func/Proc.: $ShellExecuteAndWait
* Descricao.: Executa um processo externo e aguarda o retorno
* Data......: 24/11/2009
* Por.......: Carlos H. Cantu - www.firebase.com.br
* Obs.......: Adaptada por Guilherme Vieira.
----------------------------------------------------------}
function TfrmPrincipal.ShellExecuteAndWait(cArquivoPrograma, cParametros: string; iOpcaoJanela: integer): DWORD;

  // Espera at� que o processo criado seja encerrado
  procedure WaitForExec(processHandle: THandle);
  var
    msg: TMsg;
    ret: DWORD;
  begin
    // Fica em loop esperando o processo terminar
    repeat
      ret := MsgWaitForMultipleObjects(1, { diz para aguardar }
             processHandle, { handle do processo }
             False, { "acorda" com qualquer evento }
             INFINITE, { espera o quanto for necess�rio }
             QS_PAINT or { "acorda" em mensagens de PAINT }
             QS_SENDMESSAGE { "acorda" com msgs enviadas por outras threads });

      if ret = WAIT_FAILED then exit; { se falhou, cai fora... }

      if ret = (WAIT_OBJECT_0 + 1) then
      begin
        {Recebeu uma mensagem, mas processa apenas mensagens de PAINT}
        while PeekMessage(msg, 0, WM_PAINT, WM_PAINT, PM_REMOVE) do
          DispatchMessage(msg);
      end;
    until ret = WAIT_OBJECT_0;
  end;

var
  ShellExecuteInfo: TShellExecuteInfo;
begin
  FillChar(ShellExecuteInfo, SizeOf(ShellExecuteInfo), #0);
  with ShellExecuteInfo do
  begin
    cbSize := SizeOf(TShellExecuteInfo);
    fMask := SEE_MASK_NOCLOSEPROCESS;
    Wnd := application.Handle;
    lpVerb := 'open';
    lpFile := PCHAR(cArquivoPrograma);
    lpParameters := PCHAR(cParametros);
    lpDirectory := nil;
    nShow := iOpcaoJanela;
  end;

  ShellExecuteEx(@ShellExecuteInfo);
  if ShellExecuteInfo.hProcess = 0 then
    Result := DWORD(-1) { Falhou na cria��o do processo}
  else
  begin
    // Aguarda pelo encerramento do processo
    WaitforExec(ShellExecuteInfo.hProcess);
    // Recupera o c�digo de retorno da aplica��o
    GetExitCodeProcess(ShellExecuteInfo.hProcess, Result);
    // Fecha o handle do processo para liberar os resources
    CloseHandle(ShellExecuteInfo.hProcess);
  end;
end;

procedure TfrmPrincipal.CriaBat(CaminhoFB, PATHBD: string);
var
  F: TextFile;
  DirFocus: string;
begin
  DirFocus := ExtractFilePath(PATHBD);

  if FileExists(ExtractFilePath(Application.ExeName)+'REPARA.BAT') then
    DeleteFile(ExtractFilePath(Application.ExeName)+'REPARA.BAT');

  AssignFile(F,ExtractFilePath(Application.ExeName)+'REPARA.BAT');
  Rewrite(F);

  Writeln(F,'path=' + CaminhoFB);
  Writeln(F,'');
  Writeln(F,'del temp.foc');
  Writeln(F,'del focus.ant');
  Writeln(F,'copy focus.foc temp.foc');
  Writeln(F,'ren focus.foc focus.Ant');
  Writeln(F,'');
  Writeln(F,'set isc_user=SYSDBA');
  Writeln(F,'set isc_password=masterkey');
  Writeln(F,'');
  Writeln(F,'gfix -m -i temp.foc');
  Writeln(F,'gfix -m -i temp.foc');
  Writeln(F,'gfix -m -i temp.foc');
  Writeln(F,'');
  Writeln(F,'gbak -g -b -z -v -i temp.foc back.gbk');
  Writeln(F,'gbak -g -r -z -v back.gbk focus.foc');
  Writeln(F,'');
  Writeln(F,'del back.gbk');
  Writeln(F,'del temp.foc');
  Writeln(F,'');
  Writeln(F,'cd \');
  Writeln(F,'echo *****  fim do reparar.bat ****');
  Writeln(F,'ren');
  CloseFile(F);
end;

end.
