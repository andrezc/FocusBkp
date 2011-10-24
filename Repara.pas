unit Repara;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DUtilit, IB_Components, IB_Access,
  IBODataset;

type
  TfrmRepara = class(TForm)
    grp1: TGroupBox;
    edtBinFirebird: TEdit;
    btnPathFB: TBitBtn;
    grp2: TGroupBox;
    edtPathFocus: TEdit;
    btnPathBD: TBitBtn;
    lbl1: TLabel;
    btnExecutar: TBitBtn;
    btn2: TBitBtn;
    pb1: TProgressBar;
    lblMensagem: TLabel;
    IBODatabase1: TIBODatabase;
    procedure btnExecutarClick(Sender: TObject);
  private
    procedure MostraMensagem(txt: string = '');
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRepara: TfrmRepara;

implementation


{$R *.dfm}

procedure TfrmRepara.MostraMensagem(txt: string = '');
begin
  lblMensagem.Caption := txt;
  Application.ProcessMessages;
end;


procedure TfrmRepara.btnExecutarClick(Sender: TObject);
var
  DirFocus: string;
begin
  {MostraMensagem('Aguarde Validando Informa��es...');

  if not IBODatabase1.Connected then begin
    if not Vazio(edtPathFocus.Text) then begin
      if FileExists(edtPathFocus.Text) then begin
        IBODatabase1.Database := edtPathFocus.Text;
        IBODatabase1.Connected := True;
      end;
    end
    else
    begin
      ShowMessage('N�o foi poss�vel estabelecer a conex�o com o banco de dados, verifique!');
      Exit;
    end;
  end;
  if Vazio(edtBinFirebird.Text) or Vazio(edtPathFocus.Text) then begin
    if Vazio(edtBinFirebird.Text) then
      ShowMessage('Preencha o caminho da pasta BIN do Firebird para prosseguir!');
    if Vazio(edtPathFocus.Text) then
      ShowMessage('Preencha o caminho do banco de dados FOCUS.FOC para prosseguir!');
    Exit;
  end;
  if not FileExists(edtBinFirebird.Text + '\instsvc.exe') then begin
    ShowMessage(edtBinFirebird.Text + '\instsvc.exe');
    ShowMessage('Verifique o caminho da pasta BIN do Firebird e se existe o arquivo INSTSVC.EXE dentro dela!'+pl+'Caso contr�rio verifique a instala��o do Firebird!');
    Exit;
  end;
  if not FileExists(edtPathFocus.Text) then begin
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
                 }
end;

end.
