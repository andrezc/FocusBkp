program FocusBackup;

uses
  ExceptionLog,
  Forms,
  Principal in 'Principal.pas' {frmPrincipal},
  Backup in 'Backup.pas' {frmBackup},
  Restauracao in 'Restauracao.pas' {frmRestauracao},
  Repara in 'Repara.pas' {frmRepara};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.CreateForm(TfrmBackup, frmBackup);
  Application.CreateForm(TfrmRestauracao, frmRestauracao);
  Application.CreateForm(TfrmRepara, frmRepara);
//  Application.CreateForm(TfDB, fDB);
  Application.Run;
end.
