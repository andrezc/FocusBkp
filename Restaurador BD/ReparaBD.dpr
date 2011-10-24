program ReparaBD;

uses
  ExceptionLog,
  Forms,
  UPrincipal in 'UPrincipal.pas' {frmPrincipal};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Focus Reparador';
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.Run;
end.
