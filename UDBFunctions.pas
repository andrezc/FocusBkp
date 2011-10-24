unit UDBFunctions;

interface

uses
  SysUtils, DB, DBClient, DBGrids, Graphics, Classes, IBODataset, IB_Components,
  Grids, Windows, IB_Access;

type
  TConexao = (tcServer, tcLocal, tcAutomatico, tcSuperLocal);
  TUpdateInsert = (uiUpdate, uiInsert, uiUpdateOrInsert);
  TResultVetor = array of Variant;

  TfDB = class(TDataModule)
  private
    // Private
  public
    Falhas : String;
    function  CriaCampo(ccConexao: TConexao; ccTabela, ccCampo, ccDomain: String): Boolean;
    function  TabelaTaVazia(etTabela:String; etCondicao:String; eTConexao:TConexao=tcServer):Boolean;

    procedure DestacaColunaOrdenada(Grid: TDBGrid; Rect: TRect;
              State: TGridDrawState; DataCol:Integer; Column, ColunaOrdenada: TColumn);
    procedure GridZebrado (RecNo:LongInt; Grid:TDBGrid; Rect:TRect;
              State:TGridDrawState; Column:TColumn; ColunaOrdenada:TColumn=nil;
              CorSim:TColor=clMoneyGreen; CorNao:TColor=clWhite; CorSelecionado:TColor=$00619FE4);
    procedure OrdenaDataSetGrid(var CDS: TClientDataSet; Column: TColumn; var dbgPrin: TDBGrid);
    function  SomaColunaCDS(var CDS: TClientDataSet; Coluna: String): Real;

    function  FirebirdOk(foConexao:TConexao=tcServer): Boolean;
    function  PegaVersaoDoFirebird(pvConexao:TConexao=tcServer):String;
    function  TipoDoCampo(tcTabela, tcCampo: String; tcTipoConexao:TConexao=tcServer): integer;
    function  ExcluiChavesEIndicesDoCampo(eiTabela, eiCampo:String; eiConexao:TConexao=tcServer):boolean;
    function  PegaCamposChavePrimaria(pdTabela:String; pdConexao:TConexao=tcServer):TResultVetor;
    procedure PegaTextoDoCampo(var ptTexto: TStrings; ptCampo, ptTabela, ptCondicao: String; ptConexao:TConexao=tcServer);
    function  CampoEChavePrimaria(pdTabela,pdCampo:String; pdConexao:TConexao=tcServer):boolean;
    function  TabelaDaQuery(tqQuery: TStrings): String;
    procedure CriaRegistroSeNaoExistir(crTabela, crCampo: String; crCampoValor: Variant; crCampos: array of String; crValores: array of Variant; crConexao:TConexao=tcServer);
    function  TamanhoDoCampo(tcTabela, tcCampo: String; tcTipoConexao:TConexao=tcServer): integer;
    procedure EsvaziaTabela(etTabela, etCondicao:String; eTConexao:TConexao=tcServer);
    function  ExisteTabela(ttNomeDaTabela:String; tTConexao:TConexao=tcServer): boolean;
    function  ExisteCampo(tcNomeDoCampo, tcNomeDaTabela:String; tcTipoConexao:TConexao=tcServer): boolean;
    function  TemCampoNumTQuery(tcTabela: TIBOQuery; tcNomeDoCampo: String): boolean;
    procedure PegaCampos(pcNomeDaTabela: String; var Lista: TStrings; pcConexao:TConexao=tcServer);
    function  ExisteParametro(epTabela: TIBOQuery; epParametro: String): Boolean; overload;
    function  ExisteParametro(epTabela: TIB_Query; epParametro: String): Boolean; overload;
    procedure FormataCampo(fcTipoDoCampo: word; fcCampo: TField; fcDisplayFormat: String; fcAlinhamento: TAlignment);
    procedure MontaSQLParaGravacao(msTabela: String; msIBOTabela: TIB_Query; msListaDeCampos: TStrings; msModo: TUpdateInsert);
    function  ConexoesSimultaneas(csTipoConexao:TConexao=tcServer):Integer;
    function  PegaValorDoCampo(pvCampo, pvTabela, pvCondicao: String; pvDefault:Variant; pvTipoConexao:TConexao=tcServer):Variant;
    function  PegaValoresDosCampos(pvTabela: String; pvCampos: array of String; pvDefault: array of Variant; pvCondicao: String; pvTipoConexao:TConexao=tcServer): TResultVetor;
    function  PegaValorDaQuery(pvSQL:String; pvDefault:Variant; pvTipoConexao:TConexao=tcServer):Variant;
    function  PegaValoresDaQuery(pvSQL:String; pvDefault:array of Variant; pvTipoConexao:TConexao=tcServer):TResultVetor;
    function  CampoNotNull(nnNomeDoCampo, nnNomeDaTabela:String; tTConexao:TConexao=tcServer):boolean;
    function  DefineValor(dvTabela, dvCampo: String; dvValor:Variant; dvCondicao:String; dvIncluiSeVazio:Boolean; dvTipoConexao:TConexao=tcServer): boolean;
    function  DefineValores(dvNomeDaTabela:String; dvCamposValores: Array of Variant; dvCondicao:String; dvIncluiSeVazio:Boolean; dvTipoConexao:TConexao=tcServer{; dvTabela:TIB_Query=nil; dvTransacao:TIB_Transaction=nil}): boolean;
    function  CadastraValores(cvTabela:String; cvCamposValores: array of Variant; cvTipoConexao:TConexao=tcServer): boolean;
    function  AlteraValores(avNomeDaTabela:String; avCamposValores:array of variant; avCondicao:String; avTipoConexao:TConexao=tcServer): boolean;
    function  ApagaRegistros(arNomeDaTabela, arCondicao: String; arTipoConexao:TConexao=tcServer): boolean;
    procedure AjustaGenerator(agTabela, agCampoChave, agGenerator:String; agTipoConexao:TConexao=tcServer);
    function  CriaGenerator(ggNomeDoGenerator: String; ggConexao:TConexao=tcServer): boolean;
    procedure GravaGenerator(ggNomeDoGenerator:String; ggValor:Integer; ggTipoConexao:TConexao=tcServer);
    function  ExisteRegistro(erTabela, erCampo: String; erValor: Variant; erTipoConexao:TConexao=tcServer):boolean;
    function  QuantidadeDeRegistros(qrTabela:String; qrCondicao:String; qrConexao:TConexao=tcServer):Integer;
    function  ValorMaximo(vmTabela, vmCampo, vmCondicao:String; vmConexao:TConexao=tcServer):Variant;
    function  PegaValorDoGenerator(vgGenerator:String; vgIncrementaValor:boolean=true; vgConexao:TConexao=tcServer):Integer;
    procedure ApagaTriggersEProcedures(atpConexao:TConexao=tcServer);
    function  ExisteDomain(edDomain:String; edConexao:TConexao=tcServer): Boolean;
    function  NaoTemTrigger(nttConexao:TConexao=tcServer): boolean;
    function  PegaDomain(pdTabela, pdCampo:String; pdConexao:TConexao=tcServer): String;
    procedure CreateQueryTransacao(var cqQuery:TIB_Query; var cqTransacao:TIB_Transaction; cqTipoConexao:TConexao=tcServer; cqNovaTransacao:Boolean=true);
    procedure FreeQueryTransacao(var dqQuery: TIB_Query; var dqTransacao: TIB_Transaction);
    procedure FechaTT(ftIBQuery:TIB_Query); overload;
    procedure FechaTT(ftIBQuery:TIBOQuery); overload;
    function  ExecutaSQL_EComita(esSQL: String; esParametrosValores:Array of Variant; esConexao:TConexao=tcServer): Boolean; overload;
    function  ExecutaSQL_EComita(esSQL:TStrings; esParametrosValores:Array of Variant; esConexao:TConexao=tcServer):Boolean; overload;
    function  ExecutaSQL_SemComitar(esTabela:TIB_Query; esSQL: String; esParametrosValores:Array of Variant): boolean;   overload;
    function  ExecutaSQL_SemComitar(esTabela:TIB_Query; esSQL: TStrings; esParametrosValores:Array of Variant): boolean; overload;
    function  DefineCampoNotNull(dcTabela, dcCampo: String; dcConexao:TConexao=tcServer): boolean;
    function  ExcluiIndicesDaTabela(eiTabela:String; eiConexao:TConexao=tcServer): boolean;
    function  ExisteIndiceOuChave(icNome:String; icConexao:TConexao=tcServer):boolean;
    function  RecriaUmaChavePK(rcTabela, rcCampo, rcChave:String; rcConexao:TConexao=tcServer; rcExibeMensagem:boolean=false; rcApaga:boolean=false):boolean;
    function  RecriaUmIndiceIDX(riTabela, riCampo, riIndice:String; riConexao:TConexao=tcServer; riExibeMensagem:boolean=false; riApaga:boolean=false):boolean;
    procedure AlteraDomain(adTabela, adCampo, adDomain: String; adConexao:TConexao=tcServer);
  end;

var
  fDB: TfDB;
  TaSemTrigger  : boolean;
  ResultVetor   : TResultVetor;   // Para ser usado com a função PegaValoresDosCampos e PegaValoresDaQuery

implementation

{$R *.dfm}

uses Dialogs, Variants, DUtilit, UConstVar, UComum;

procedure TfDB.CreateQueryTransacao(var cqQuery:TIB_Query; var cqTransacao:TIB_Transaction; cqTipoConexao:TConexao=tcServer; cqNovaTransacao:Boolean=true);
//var
//  NovaTransacao : boolean;
begin
  try
//    NovaTransacao := cqTransacao = nil;
    cqQuery := TIB_Query.Create(self);
    case cqTipoConexao of
      tcServer     : cqQuery.IB_Connection := CM.IBTabela.IB_Connection;
      tcLocal      : cqQuery.IB_Connection := CM.IBTabelaLocal.IB_Connection;
      tcSuperLocal : begin
                       if not CM.IBDatabaseSuperLocal.Connected then
                       begin
                         CM.IBDatabaseSuperLocal.DatabaseName := PathComBarra(DiretorioDoExecutavel)+ArqBancoLocal;
                         CM.IBDatabaseSuperLocal.Connect;
                       end;
                       cqQuery.IB_Connection := CM.IBDatabaseSuperLocal;
                     end;
    end;
    if cqNovaTransacao then cqTransacao := TIB_Transaction.Create(self);
    cqTransacao.IB_Connection := cqQuery.IB_Connection;
    cqQuery.IB_Transaction := cqTransacao;
    cqQuery.Active := false;
  except
  end;
end;

procedure TfDB.FreeQueryTransacao(var dqQuery:TIB_Query; var dqTransacao:TIB_Transaction);
begin
  if dqQuery     <> nil then try dqQuery.Free;     except end;
  if dqTransacao <> nil then try dqTransacao.Free; except end;
end;

procedure TfDB.FechaTT(ftIBQuery:TIB_Query);
begin
  try
    if ftIBQuery.IB_Transaction.InTransaction then ftIBQuery.IB_Transaction.Commit;
    ftIBQuery.Active := false;
  except
  end;
end;

procedure TfDB.FechaTT(ftIBQuery:TIBOQuery);
begin
  try
    if ftIBQuery.IB_Transaction.InTransaction then ftIBQuery.IB_Transaction.Commit;
    ftIBQuery.Active := false;
  except
  end;
end;

function TfDB.ExecutaSQL_SemComitar(esTabela:TIB_Query; esSQL:String; esParametrosValores:Array of Variant):boolean;
var
  i : Word;
begin
  try
    esTabela.Active := false;
    esTabela.SQL.Clear;
    esTabela.SQL.Add(esSQL);
    if not esTabela.Prepared then esTabela.Prepare;
    if length(esParametrosValores) > 0 then
    begin
      for i := 0 to Length(esParametrosValores)-1 do
      begin
        if i mod 2 = 0 then  // Se for par, associa o parâmetro
          esTabela.ParamByName(esParametrosValores[i]).Value := esParametrosValores[i+1];
      end;
    end;
    esTabela.ExecSQL;
    result := true;
  except
    result := false;
  end;
end;

function TfDB.ExecutaSQL_SemComitar(esTabela:TIB_Query; esSQL:TStrings; esParametrosValores:Array of Variant):boolean;
var
  i : Word;
begin
  try
    esTabela.Active := false;
    esTabela.SQL.Clear;
    esTabela.SQL.Assign(esSQL);
    if not esTabela.Prepared then esTabela.Prepare;
    if length(esParametrosValores) > 0 then
    begin
      for i := 0 to Length(esParametrosValores)-1 do
      begin
        if i mod 2 = 0 then  // Se for par, associa o parâmetro
          esTabela.ParamByName(esParametrosValores[i]).Value := esParametrosValores[i+1];
      end;
    end;
    esTabela.ExecSQL;
    result := true;
  except
    result := false;
  end;
end;

function TfDB.FirebirdOk(foConexao:TConexao=tcServer):Boolean;
const
  VersaoMinima = '2.5.0';
begin
  try
    if Vazio(VersaoDoFirebird) then VersaoDoFirebird := PegaVersaoDoFirebird(tcServer);
  except
  end;

  result := ( (StrToIntZ(StringPos(VersaoDoFirebird, '.', 1), 0) >= (StrToIntZ(StringPos(VersaoMinima, '.', 1), 0)) ) and
              (StrToIntZ(StringPos(VersaoDoFirebird, '.', 2), 0) >= (StrToIntZ(StringPos(VersaoMinima, '.', 2), 0)) ) and
              (StrToIntZ(StringPos(VersaoDoFirebird, '.', 3), 0) >= (StrToIntZ(StringPos(VersaoMinima, '.', 3), 0)) )  );

//result := False;
  if not result then
  begin
    CM.MensagemDeAtencao('A T E N Ç Ã O !' + PL+PL+
                         'É necessário a instalação do Firebird 2.5 ou superior para esta versão do FOCUS.'+PL+PL+'Baixe o Firebird em '+PL+'http://www.firebirdsql.org');
    exit;
  end;
end;

function TfDB.ExecutaSQL_EComita(esSQL:String; esParametrosValores:Array of Variant; esConexao:TConexao=tcServer):Boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
  i : Word;
begin
  result := false;
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, esConexao);
    try
      EstaTabela.IB_Transaction.StartTransaction;
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.add(esSQL);
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      if length(esParametrosValores) > 0 then
      begin
        for i := 0 to Length(esParametrosValores)-1 do
        begin
          if i mod 2 = 0 then  // Se for par, associa o parâmetro
            EstaTabela.ParamByName(esParametrosValores[i]).Value := esParametrosValores[i+1];
        end;
      end;
      EstaTabela.ExecSQL;
      EstaTabela.IB_Transaction.Commit;
      result := true;
    except
      result := false;
      EstaTabela.IB_Transaction.Rollback;
    end;
  finally
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.ExecutaSQL_EComita(esSQL:TStrings; esParametrosValores:Array of Variant; esConexao:TConexao=tcServer):Boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
  i : Word;
begin
  result := false;
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, esConexao);
    try
      EstaTabela.IB_Transaction.StartTransaction;
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Assign(esSQL);
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      if length(esParametrosValores) > 0 then
      begin
        for i := 0 to Length(esParametrosValores)-1 do
        begin
          if i mod 2 = 0 then  // Se for par, associa o parâmetro
            EstaTabela.ParamByName(esParametrosValores[i]).Value := esParametrosValores[i+1];
        end;
      end;
      EstaTabela.ExecSQL;
      EstaTabela.IB_Transaction.Commit;
      result := true;
    except
      EstaTabela.IB_Transaction.Rollback;
      result := false;
    end;
  finally
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.ExisteIndiceOuChave(icNome:String; icConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, icConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('select rdb$index_name from rdb$indices where rdb$index_name = :INDICE');
    if not EstaTabela.Prepared then EstaTabela.Prepare;
    EstaTabela.ParamByName('INDICE').AsString := icNome;
    EstaTabela.Active := true;
    result := not EstaTabela.IsEmpty;
  finally
    FechaTT(EstaTabela);
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.DefineCampoNotNull(dcTabela, dcCampo:String; dcConexao:TConexao=tcServer): boolean;
var
  EsteCampo : String;
  i, QuantCampos : word;
  falhou : boolean;
begin
  QuantCampos := ContaCaracteres(dcCampo, ',')+1;
  if QuantCampos = 1 then
  begin
    result := ExecutaSQL_EComita('update RDB$RELATION_FIELDS set RDB$NULL_FLAG = 1 where (RDB$FIELD_NAME = '+QuotedStr(dcCampo)+') and (RDB$RELATION_NAME = '+QuotedStr(dcTabela)+')', [], dcConexao);
  end

  else begin
    result := false;
    falhou := false;
    for i := 1 to QuantCampos do
    begin
      EsteCampo := TiraEspacos(StringPos(dcCampo, ',', i));
      if not ExecutaSQL_EComita('update RDB$RELATION_FIELDS set RDB$NULL_FLAG = 1 where (RDB$FIELD_NAME = '+QuotedStr(EsteCampo)+') and (RDB$RELATION_NAME = '+QuotedStr(dcTabela)+')', [], dcConexao) then falhou := true;
    end;
    result := not falhou;
  end;
end;

function TfDB.RecriaUmaChavePK(rcTabela, rcCampo, rcChave:String; rcConexao:TConexao=tcServer; rcExibeMensagem:boolean=false; rcApaga:boolean=false):boolean;
var
  OldChave : String;
begin
  result := false;
  try
    if rcExibeMensagem then CM.AbreAviso('Recriando Chave Primária '+rcChave);
    if rcApaga then ExecutaSQL_EComita('ALTER TABLE '+rcTabela+' DROP CONSTRAINT '+rcChave, [], rcConexao);
    DefineCampoNotNull(rcTabela, rcCampo);

    if not ExecutaSQL_EComita('ALTER TABLE '+rcTabela+' ADD CONSTRAINT '+rcChave+' PRIMARY KEY ('+rcCampo+')', [], rcConexao) then Raise Exception.Create('Não foi possível Criar Chave PK');
    result := true;
  except
    result := false;
    Falhas := Falhas + rcChave + PL;
  end;
end;

function TfDB.RecriaUmIndiceIDX(riTabela, riCampo, riIndice:String; riConexao:TConexao=tcServer; riExibeMensagem:boolean=false; riApaga:boolean=false):boolean;
begin
  result := false;
  try
    if riExibeMensagem then CM.AbreAviso('Recriando Índice '+riIndice);
    if riApaga then ExecutaSQL_EComita('DROP INDEX '+riIndice, [], riConexao);

    if not ExecutaSQL_EComita('CREATE INDEX  '+riIndice+' ON '+riTabela+' ('+riCampo+')', [], riConexao) then Raise Exception.Create('Não foi possível Criar Índice');
    result := true;
  except
    result := false;
    Falhas := Falhas + riIndice + PL;
  end;
end;

function TfDB.ExcluiIndicesDaTabela(eiTabela:String; eiConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  result := false;
  try
    try
      CreateQueryTransacao(EstaTabela, EstaTransacao, eiConexao);
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('select rdb$indices.RDB$INDEX_NAME as IND from rdb$indices where coalesce(rdb$indices.rdb$unique_flag, 0) = 0 and RDB$RELATION_NAME = '+QuotedStr(eiTabela));
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.Active := true;
      while not EstaTabela.eof do
      begin
        try ExecutaSQL_EComita('DROP INDEX '+EstaTabela.FieldByName('IND').AsString, []); except end;
        EstaTabela.Next;
      end;
      result := true;
    except
      result := false;
    end;
  finally
    FechaTT(EstaTabela);
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.ExcluiChavesEIndicesDoCampo(eiTabela, eiCampo:String; eiConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  result := false;
  try
    try
      CreateQueryTransacao(EstaTabela, EstaTransacao, eiConexao);

      EstaTabela.Active := false;  // Pega chaves primária...
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('select S.RDB$INDEX_NAME CHAVE from RDB$INDEX_SEGMENTS S');
      EstaTabela.SQL.Add('join RDB$INDICES I on I.RDB$INDEX_NAME = S.RDB$INDEX_NAME');
      EstaTabela.SQL.Add('where I.RDB$RELATION_NAME = :TABELA');
      EstaTabela.SQL.Add('and S.RDB$FIELD_NAME = :CAMPO');
      EstaTabela.SQL.Add('and I.RDB$UNIQUE_FLAG = 1');
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ParamByName('TABELA').AsString := eiTabela;
      EstaTabela.ParamByName('CAMPO').AsString := eiCampo;
      EstaTabela.Active := true;
      while not EstaTabela.Eof do    // Detona as Chaves Primarias
      begin
        ExecutaSQL_EComita('ALTER TABLE '+eiTabela+' DROP CONSTRAINT '+EstaTabela.FieldByName('CHAVE').AsString, [], eiConexao);
        EstaTabela.Next;
      end;

      EstaTabela.Active := false;  // Pega índices...
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('select S.RDB$INDEX_NAME INDICE from RDB$INDEX_SEGMENTS S');
      EstaTabela.SQL.Add('join RDB$INDICES I on I.RDB$INDEX_NAME = S.RDB$INDEX_NAME');
      EstaTabela.SQL.Add('where I.RDB$RELATION_NAME = :TABELA');
      EstaTabela.SQL.Add('and S.RDB$FIELD_NAME = :CAMPO');
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ParamByName('TABELA').AsString := eiTabela;
      EstaTabela.ParamByName('CAMPO').AsString := eiCampo;
      EstaTabela.Active := true;
      while not EstaTabela.Eof do    // Detona os Índices
      begin
        ExecutaSQL_EComita('DROP INDEX '+EstaTabela.FieldByName('INDICE').AsString, [], eiConexao);
        EstaTabela.Next;
      end;
      FechaTT(EstaTabela);

      result := true;
    except
      result := false;
    end;
  finally
    FechaTT(EstaTabela);
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

procedure TfDB.AlteraDomain(adTabela, adCampo, adDomain: String; adConexao:TConexao=tcServer);
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
  Chave        : String;
  CamposChave  : String;
  ChaveApagada : boolean;
  S : String;
begin
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, adConexao);
    ChaveApagada := false;
    CamposChave := '';
    Chave := '';
    S := '';
    try
      EstaTabela.Active := false;  // Vê se o campo é chave primária...
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('select S.RDB$INDEX_NAME CHAVE from RDB$INDEX_SEGMENTS S');
      EstaTabela.SQL.Add('join RDB$INDICES I on I.RDB$INDEX_NAME = S.RDB$INDEX_NAME');
      EstaTabela.SQL.Add('where I.RDB$RELATION_NAME = :TABELA');
      EstaTabela.SQL.Add('and S.RDB$FIELD_NAME = :CAMPO');
      EstaTabela.SQL.Add('and I.RDB$UNIQUE_FLAG = 1');
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ParamByName('TABELA').AsString := adTabela;
      EstaTabela.ParamByName('CAMPO').AsString := adCampo;
      EstaTabela.Active := true;
      if not EstaTabela.IsEmpty then
        Chave := EstaTabela.FieldByName('CHAVE').AsString;
      FechaTT(EstaTabela);

      if not Vazio(Chave) then     // Se o campo é chave primária, pega os campos desta chave
      begin
        EstaTabela.Active := false;
        EstaTabela.SQL.Clear;
        EstaTabela.SQL.Add('select S.RDB$FIELD_NAME CAMPO from RDB$INDEX_SEGMENTS S');
        EstaTabela.SQL.Add('join RDB$INDICES I on I.RDB$INDEX_NAME = S.RDB$INDEX_NAME');
        EstaTabela.SQL.Add('where S.RDB$INDEX_NAME = :CHAVE');
        if not EstaTabela.Prepared then EstaTabela.Prepare;
        EstaTabela.ParamByName('CHAVE').AsString := Chave;
        EstaTabela.Active := true;
        if not EstaTabela.IsEmpty then
        begin
          CamposChave := EstaTabela.FieldByName('CAMPO').AsString;
          EstaTabela.Next;
          while not EstaTabela.Eof do
          begin
            CamposChave := CamposChave + ','+ EstaTabela.FieldByName('CAMPO').AsString;
            EstaTabela.Next;
          end;
        end;
        FechaTT(EstaTabela);
      end;

      if not Vazio(Chave) then     // Se o campo é chave primária, exclui a chave...
      begin
        EstaTabela.IB_Connection.Disconnect;   // Tive que colocar esse troço aí porque tava dando erro:
        EstaTabela.IB_Connection.Connect;      //

        if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
        EstaTabela.Active := false;
        EstaTabela.SQL.Clear;
        EstaTabela.SQL.Add('ALTER TABLE '+adTabela+' DROP CONSTRAINT '+Chave);
        if not EstaTabela.Prepared then EstaTabela.Prepare;
        EstaTabela.ExecSQL;
        EstaTabela.IB_Transaction.Commit;
        ChaveApagada := true;
      end;

      if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;                    // Altera o Domain
      EstaTabela.SQL.Add('update RDB$RELATION_FIELDS set');
      EstaTabela.SQL.Add('RDB$FIELD_SOURCE = '+QuotedStr(adDomain));
      EstaTabela.SQL.Add('where ((RDB$FIELD_NAME = '+QuotedStr(adCampo)+') and');
      EstaTabela.SQL.Add('(RDB$RELATION_NAME = '+QuotedStr(adTabela)+'))');
      EstaTabela.Prepare;
      EstaTabela.ExecSQL;
      EstaTabela.IB_Transaction.Commit;

      if not Vazio(Chave) then
      begin                      // Recria a chave se o campo é chave primária...
        if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
        EstaTabela.Active := false;
        EstaTabela.SQL.Clear;
        EstaTabela.SQL.Add('alter table '+adTabela+' add constraint PK_'+adTabela+' primary key ('+CamposChave+')');
        if not EstaTabela.Prepared then EstaTabela.Prepare;
        EstaTabela.ExecSQL;
        EstaTabela.IB_Transaction.Commit;
      end;

      EstaTabela.IB_Transaction.Commit;
    except
      if ChaveApagada then
        S := PL+PL+'ATENÇÃO!!!'+PL+'A chave primária '+chave+' foi excluída!';

      MessageBox(0, PChar('Não foi possível Alterar o DOMAIN!!!'+PL+
                 'Tabela: '+adTabela+PL+
                 'Campo: '+adCampo+PL+
                 'Domain: '+adDomain+S), 'ATENÇÃO', MB_ICONWARNING or MB_OK);
    end;
  finally
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.PegaDomain(pdTabela, pdCampo:String; pdConexao:TConexao=tcServer):String;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, pdConexao);
    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    try
      EstaTabela.SQL.Add('select rdb$field_source D from rdb$relation_fields');
      EstaTabela.SQL.Add('where rdb$Field_Name = :CAMPO and rdb$Relation_Name = :TABELA');
      EstaTabela.ParamByName('CAMPO').AsString := pdCAMPO;
      EstaTabela.ParamByName('TABELA').AsString := pdTABELA;
      EstaTabela.Prepare;
      EstaTabela.Active := true;
      result := EstaTabela.FieldByName('D').AsString;
      FechaTT(EstaTabela);
    except
    end;
  finally
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.ExisteDomain(edDomain:String; edConexao:TConexao=tcServer):Boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, edConexao);
    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    try
      EstaTabela.SQL.Add('select rdb$field_name from RDB$FIELDS where rdb$field_name = :D');
      EstaTabela.ParamByName('D').AsString := edDomain;
      EstaTabela.Prepare;
      EstaTabela.Active := true;
      result := not EstaTabela.IsEmpty;
      FechaTT(EstaTabela);
    except
    end;
  finally
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

procedure TfDB.ApagaTriggersEProcedures(atpConexao:TConexao=tcServer);
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
  IBTabela2 : TIB_Query;
  Transacao2 : TIB_Transaction;
begin
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, atpConexao);
    CreateQueryTransacao(IBTabela2, Transacao2, atpConexao);
    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;

    // TRIGGERS
    IBTabela2.Active := false;
    IBTabela2.SQL.Clear;
    IBTabela2.SQL.Add('select RDB$TRIGGER_NAME as NOME from RDB$TRIGGERS');
    IBTabela2.SQL.Add('where RDB$TRIGGER_NAME not like ''RDB$%''');
    IBTabela2.Active := true;
    IBTabela2.First;

    if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
    try
      while not IBTabela2.eof do
      begin
        if copy(IBTabela2.FieldByName('NOME').AsString, 1, 4) <> 'RDB$' then
        begin
          EstaTabela.Active := false;
          EstaTabela.SQL.Clear;
          EstaTabela.SQL.Add('drop TRIGGER '+IBTabela2.FieldByName('NOME').AsString);
          EstaTabela.ExecSQL;
        end;
        IBTabela2.Next;
      end;
      fDB.FechaTT(IBTabela2);
      EstaTabela.IB_Transaction.Commit;
    except
      EstaTabela.IB_Transaction.RollBack;
      CM.MensagemDeErro('Erro ao excluir as Triggers');
    end;

    // PROCEDURES
    IBTabela2.Active := false;
    IBTabela2.SQL.Clear;
    IBTabela2.SQL.Add('select RDB$PROCEDURE_NAME as NOME from RDB$PROCEDURES');
    IBTabela2.Active := true;
    IBTabela2.First;

    if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
    try
      while not IBTabela2.eof do
      begin
        EstaTabela.Active := false;
        EstaTabela.SQL.Clear;
        EstaTabela.SQL.Add('drop PROCEDURE '+IBTabela2.FieldByName('NOME').AsString);
        EstaTabela.ExecSQL;
        IBTabela2.Next;
      end;
      fDB.FechaTT(IBTabela2);
      EstaTabela.IB_Transaction.Commit;
      TaSemTrigger := true;
    except
      EstaTabela.IB_Transaction.RollBack;
      CM.MensagemDeErro('Erro ao excluir as Procedures');
    end;
  finally
    FreeQueryTransacao(EstaTabela, EstaTransacao);
    FreeQueryTransacao(IBTabela2, Transacao2);
  end;
end;

function TfDB.NaoTemTrigger(nttConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, nttConexao);
    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('select RDB$TRIGGER_NAME as NOME from RDB$TRIGGERS');
    EstaTabela.SQL.Add('where RDB$TRIGGER_NAME not like ''RDB$%''');
    EstaTabela.Active := true;
    NaoTemTrigger := EstaTabela.IsEmpty;
    fDB.FechaTT(EstaTabela);
  finally
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.PegaValorDoGenerator(vgGenerator:String; vgIncrementaValor:boolean=true; vgConexao:TConexao=tcServer):Integer;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
  i : Integer;
begin
  try
    try
      fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, vgConexao);

      if vgIncrementaValor then i := 1
                           else i := 0;

      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('select GEN_ID('+vgGenerator+', '+intToStr(i)+') as GEN from rdb$database');
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.Active := true;
      result := EstaTabela.FieldByName('GEN').AsInteger;
    except
      result := 0;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end
end;

function TfDB.ExisteRegistro(erTabela, erCampo: String; erValor: Variant; erTipoConexao:TConexao=tcServer):boolean;
var
  Existe: boolean;
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, erTipoConexao);

    Existe := false;
    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add ('SELECT 1 FROM '+UpperCase(erTabela)+' WHERE '+UpperCase(erCampo)+' = :VALOR');
    if not EstaTabela.Prepared then EstaTabela.Prepare;
    EstaTabela.ParamByName('VALOR').Value := erValor;
    EstaTabela.Active := true;
    try
      Existe := not EstaTabela.EOF;
    finally
      EstaTabela.Active := false;
    end;
    result := Existe;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.QuantidadeDeRegistros(qrTabela:String; qrCondicao:String; qrConexao:TConexao=tcServer):Integer;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, qrConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('SELECT count(*) FROM '+qrTabela);
    if not Vazio(qrCondicao) then
      EstaTabela.SQL.Add('where '+qrCondicao);
    if not EstaTabela.Prepared then EstaTabela.Prepare;
    EstaTabela.Active := true;
    result := EstaTabela.FieldByname('COUNT').AsInteger;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.ValorMaximo(vmTabela, vmCampo, vmCondicao:String; vmConexao:TConexao=tcServer):Variant;
var
  EstaVMTabela : TIB_Query;
  EstaVMTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaVMTabela, EstaVMTransacao, vmConexao);

    EstaVMTabela.Active := false;
    EstaVMTabela.SQL.Clear;
    EstaVMTabela.SQL.Add('SELECT coalesce(max('+vmCampo+'), 0) as MAIOR FROM '+vmTabela);
    if not Vazio(vmCondicao) then
      EstaVMTabela.SQL.Add('where '+vmCondicao);

    CM.SalvaSQL(EstaVMTabela.SQL);
    if not EstaVMTabela.Prepared then EstaVMTabela.Prepare;
    EstaVMTabela.Active := true;
    result := EstaVMTabela.FieldByname('MAIOR').Value;
  finally
    fDB.FechaTT(EstaVMTabela);
    fDB.FreeQueryTransacao(EstaVMTabela, EstaVMTransacao);
  end;
end;

procedure TfDB.AjustaGenerator(agTabela, agCampoChave, agGenerator:String; agTipoConexao:TConexao=tcServer);
var
  Max : Integer;
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, agTipoConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('Select Max('+agCampoChave+') as MAXIMO from '+agTabela);
    if not EstaTabela.Prepared then EstaTabela.Prepare;
    EstaTabela.Active := true;
    if EstaTabela.IsEmpty then Max := 0
                        else Max := EstaTabela.FieldByName('MAXIMO').AsInteger;
    EstaTabela.Active := false;

    if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('set generator '+agGenerator+' to '+IntToStr(Max));
    if not EstaTabela.Prepared then EstaTabela.Prepare;
    EstaTabela.ExecSQL;
    EstaTabela.IB_Transaction.Commit;
    EstaTabela.Active := false;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

procedure TfDB.GravaGenerator(ggNomeDoGenerator:String; ggValor:Integer; ggTipoConexao:TConexao=tcServer);
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, ggTipoConexao);
    try
      if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('set generator '+ggNomeDoGenerator+' to '+IntToStr(ggValor));
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ExecSQL;
      EstaTabela.IB_Transaction.Commit;
      EstaTabela.Active := false;
    except
      EstaTabela.Active := false;
    end;
  finally
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.CriaGenerator(ggNomeDoGenerator:String; ggConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  result := false;
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, ggConexao);
    try
      if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('create sequence '+ggNomeDoGenerator);
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ExecSQL;
      EstaTabela.IB_Transaction.Commit;
      EstaTabela.Active := false;
      result := true;
    except
      result := false;
      EstaTabela.Active := false;
    end;
  finally
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.DefineValor(dvTabela, dvCampo:String; dvValor:Variant; dvCondicao:String; dvIncluiSeVazio:Boolean; dvTipoConexao:TConexao=tcServer): boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, dvTipoConexao);

    if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
    try
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('UPDATE '+Uppercase(dvTabela)+' SET '+Uppercase(dvCampo)+' = :VALOR');
      if not Vazio(dvCondicao) then EstaTabela.SQL.Add('Where ' + dvCondicao);
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ParamByName('VALOR').Value := dvValor;
      EstaTabela.ExecSQL;

      if (EstaTabela.RowsAffected = 0) and (dvIncluiSeVazio) then
      begin
        EstaTabela.Active := false;
        EstaTabela.SQL.Clear;
        EstaTabela.SQL.Add('INSERT INTO '+Uppercase(dvTabela)+' ('+dvCampo+') values (:VALOR)');
        if not EstaTabela.Prepared then EstaTabela.Prepare;
        EstaTabela.ParamByName('VALOR').Value := dvValor;
        EstaTabela.ExecSQL;
      end;

      EstaTabela.IB_Transaction.Commit;
      result := true;
    except
      on E: exception do begin
        result := false;
        UltimoErro := E.Message;
        CM.LogDeErros(E.Message);
        EstaTabela.IB_Transaction.Rollback;
      end;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.DefineValores(dvNomeDaTabela:String; dvCamposValores: Array of variant; dvCondicao:String; dvIncluiSeVazio:Boolean; dvTipoConexao:TConexao=tcServer): boolean;
var                                                 //  Posições Pares   = CAMPOS
  i : integer;                                      //  Posições Ímpares = VALORES
begin
  result := AlteraValores(dvNomeDaTabela, dvCamposValores, dvCondicao, dvTipoConexao);
  if (not result) and (dvIncluiSeVazio) then
    result := CadastraValores(dvNomeDaTabela, dvCamposValores, dvTipoConexao);
end;

function TfDB.AlteraValores(avNomeDaTabela:String; avCamposValores: Array of variant; avCondicao:String; avTipoConexao:TConexao=tcServer): boolean;
var
  i : integer;                        //  Posições Pares   = CAMPOS
  avTabela    : TIB_Query;            //  Posições Ímpares = VALORES
  avTransacao : TIB_Transaction;

  function ConteudoDoCampo(ccCampo:variant):boolean;
  begin
    result := (VarType(ccCampo) = varString) and (length(VarToStr(ccCampo))>1) and (VarToStr(ccCampo)[1] = '@');  // Se é string e tem um arroba na frente...
  end;

begin
  try
    fDB.CreateQueryTransacao(avTabela, avTransacao, avTipoConexao);

    if not avTransacao.InTransaction then avTransacao.StartTransaction;
    try
      avTabela.Active := false;
      avTabela.SQL.Clear;

      // Tenta Atualizar assumindo que o registro exista.
      avTabela.SQL.Add('UPDATE '+Uppercase(avNomeDaTabela)+' SET ');

      if ConteudoDoCampo(avCamposValores[1]) then
        avTabela.SQL.Add(Uppercase(avCamposValores[0])+' = '+TiraCaracteres(avCamposValores[1],['@']))  // recebe o conteúdo do campo mencionado
      else                                                                                              // Senão
        avTabela.SQL.Add(Uppercase(avCamposValores[0])+' = :'+Uppercase(avCamposValores[0]));           // recebe o Valor colocado
      for i := 2 to Length(avCamposValores)-1 do
      begin
        if i mod 2 = 0 then
        begin
          if ConteudoDoCampo(avCamposValores[i+1]) then
            avTabela.SQL.Add(', '+Uppercase(avCamposValores[i])+' = '+TiraCaracteres(avCamposValores[i+1],['@']))  // recebe o conteúdo do campo mencionado
          else                                                                                                // Senão
            avTabela.SQL.Add(', '+Uppercase(avCamposValores[i])+' = :'+Uppercase(avCamposValores[i]));        // recebe o Valor colocado
        end;
      end;
      if not Vazio(avCondicao) then avTabela.SQL.Add('Where ' + avCondicao);
      CM.SalvaSQL(avTabela.SQL);
      if not avTabela.prepared then avTabela.prepare;
      for i := 1 to Length(avCamposValores)-1 do
      begin
        if i mod 2 <> 0 then
        begin
          if not ConteudoDoCampo(avCamposValores[i]) then
            avTabela.ParamByName(Uppercase(VarToStr(avCamposValores[i-1]))).Value := avCamposValores[i];
        end;
      end;
      avTabela.ExecSQL;
      result := avTabela.RowsAffected > 0;
      avTabela.IB_Transaction.Commit;
    except
      on E: exception do begin
        result := false;
        UltimoErro := E.Message;
        CM.LogDeErros(E.Message);
        avTabela.IB_Transaction.Rollback;
      end;
    end;
  finally
    fDB.FreeQueryTransacao(avTabela, avTransacao);
  end;
end;

function TfDB.CadastraValores(cvTabela:String; cvCamposValores: array of Variant; cvTipoConexao:TConexao=tcServer): boolean;
var
  i : integer;                                   //  Posições Pares   = CAMPOS
  EstaTabela : TIB_Query;                        //  Posições Ímpares = VALORES
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, cvTipoConexao);

    if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
    try
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;

      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('INSERT INTO '+Uppercase(cvTabela)+' (');
      EstaTabela.SQL.Add(Uppercase(cvCamposValores[0]));
      for i := 2 to Length(cvCamposValores)-1 do
        if i mod 2 = 0 then
          EstaTabela.SQL.Add(', '+Uppercase(cvCamposValores[i]));
      EstaTabela.SQL.Add(')VALUES (');
      EstaTabela.SQL.Add(':'+Uppercase(cvCamposValores[0]));
      for i := 2 to Length(cvCamposValores)-1 do
        if i mod 2 = 0 then
          EstaTabela.SQL.Add(', :'+Uppercase(cvCamposValores[i]));
      EstaTabela.SQL.Add(')');
      CM.SalvaSQL(EstaTabela.SQL);
      if not EstaTabela.prepared then EstaTabela.prepare;

      for i := 1 to Length(cvCamposValores)-1 do
      begin
        if i mod 2 <> 0 then
          EstaTabela.ParamByName(Uppercase(VarToStr(cvCamposValores[i-1]))).Value := cvCamposValores[i];
      end;

      EstaTabela.ExecSQL;
      EstaTabela.IB_Transaction.Commit;
      result := true;
    except
      on E: exception do begin
        result := false;
        UltimoErro := E.Message;
        CM.LogDeErros(E.Message);
        EstaTabela.IB_Transaction.Rollback;
      end;
    end;
  finally
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.ApagaRegistros(arNomeDaTabela:String; arCondicao:String; arTipoConexao:TConexao=tcServer): boolean;
var
  arTabela    : TIB_Query;
  arTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(arTabela, arTransacao, arTipoConexao);

    if not arTransacao.InTransaction then arTransacao.StartTransaction;
    try
      arTabela.Active := false;
      arTabela.SQL.Clear;

      arTabela.SQL.Add('delete from '+Uppercase(arNomeDaTabela));
      if not Vazio(arCondicao) then arTabela.SQL.Add('Where ' + arCondicao);
      CM.SalvaSQL(arTabela.SQL);
      if not arTabela.prepared then arTabela.prepare;
      arTabela.ExecSQL;

      result := true;
    except
      on E: exception do begin
        result := false;
        UltimoErro := E.Message;
        CM.LogDeErros(E.Message);
        arTransacao.Rollback;
      end;
    end;
  finally
    fDB.FreeQueryTransacao(arTabela, arTransacao);
  end;
end;

function TfDB.CampoNotNull(nnNomeDoCampo, nnNomeDaTabela:String; tTConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, tTConexao);

    try
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('update RDB$RELATION_FIELDS set RDB$NULL_FLAG = 1');
      EstaTabela.SQL.Add('where (RDB$FIELD_NAME = '+QuotedStr(nnNomeDoCampo)+' and');
      EstaTabela.SQL.Add('(RDB$RELATION_NAME = '+QuotedStr(nnNomeDaTabela));
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ExecSQL;
      EstaTabela.IB_Transaction.Commit;
      result := true;
    except
      EstaTabela.IB_Transaction.Rollback;
      result := false;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.PegaValorDaQuery(pvSQL:String; pvDefault:Variant; pvTipoConexao:TConexao=tcServer):Variant;
var
  TaVazio : boolean;
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, pvTipoConexao);

    CM.TemConexaoServidorFOCUS(false);
    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add(pvSQL);
    try
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.Active := true;
      result := EstaTabela.Fields[0].Value;
      TaVazio := EstaTabela.Fields[0].IsNull;
    except
      TaVazio := true;
    end;
    if (TaVazio) then result := pvDefault;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.PegaValoresDaQuery(pvSQL:String; pvDefault:array of Variant; pvTipoConexao:TConexao=tcServer):TResultVetor;
var
  i : word;
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, pvTipoConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add(pvSQL);
    try
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.Active := true;
      setLength(result, 0);
      for i := 0 to EstaTabela.FieldCount-1 do
      begin
        setLength(result, length(result)+1);
        if EstaTabela.Fields[i].IsNull then
          result[length(result)-1] := pvDefault[i]
        else
          result[length(result)-1] := EstaTabela.Fields[i].Value;
      end;
    except
      setLength(result, 0);
      for i := 0 to EstaTabela.FieldCount-1 do
      begin
        setLength(result, length(result)+1);
        result[length(result)-1] := pvDefault[i];
      end;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.PegaValorDoCampo(pvCampo, pvTabela, pvCondicao: String; pvDefault:Variant; pvTipoConexao:TConexao=tcServer):Variant;
var
  TaVazio : boolean;
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, pvTipoConexao);

    CM.TemConexaoServidorFOCUS(false);
    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('Select '+pvCampo+' as RESULTADO from '+pvTabela);
    if not vazio(pvCondicao) then
      EstaTabela.SQL.Add('where '+pvCondicao);
    try
      EstaTabela.Active := true;
      result := EstaTabela.FieldByName('RESULTADO').Value;
//      result := EstaTabela.FieldByName('RESULTADO').AsString;
      TaVazio := EstaTabela.FieldByName('RESULTADO').IsNull;
    except
      TaVazio := true;
    end;
    if (TaVazio) then result := pvDefault;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.PegaValoresDosCampos(pvTabela:String; pvCampos:Array of String; pvDefault:array of Variant; pvCondicao:String; pvTipoConexao:TConexao=tcServer):TResultVetor;
var
  i : word;
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, pvTipoConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('Select '+pvCampos[0]+' as CAMPO1');
    for i := 1 to length(pvCampos)-1 do
      EstaTabela.SQL.Add(', '+pvCampos[i]+' as CAMPO'+inttostr(i+1));
    EstaTabela.SQL.Add('from '+pvTabela);
    if not vazio(pvCondicao) then
      EstaTabela.SQL.Add('where '+pvCondicao);
    try
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.Active := true;
      setLength(result, 0);
      for i := 0 to length(pvCampos)-1 do
      begin
        setLength(result, length(result)+1);
        if EstaTabela.FieldByName('CAMPO'+inttostr(i+1)).IsNull then
          result[length(result)-1] := pvDefault[i]
        else
          result[length(result)-1] := EstaTabela.FieldByName('CAMPO'+inttostr(i+1)).Value;
      end;
    except
      setLength(result, 0);
      for i := 0 to length(pvCampos)-1 do
      begin
        setLength(result, length(result)+1);
        result[length(result)-1] := pvDefault[i];
      end;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.ConexoesSimultaneas(csTipoConexao:TConexao=tcServer):Integer;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, csTipoConexao);
    result := EstaTabela.IB_Connection.Users.Count;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

procedure TfDB.MontaSQLParaGravacao(msTabela:String; msIBOTabela:TIB_Query; msListaDeCampos:TStrings; msModo:TUpdateInsert);
var
  i : integer;
begin
  msIBOTabela.SQL.Clear;
  case msModo of
    uiInsert, uiUpdateOrInsert:
    begin
      case msModo of
        uiInsert         : msIBOTabela.SQL.Add('Insert Into '+msTabela+' ('+msListaDeCampos[0]);
        uiUpdateOrInsert : msIBOTabela.SQL.Add('Update or Insert Into '+msTabela+' ('+msListaDeCampos[0]);
      end;
      for i := 1 to msListaDeCampos.Count -1 do
        msIBOTabela.SQL.Add(', '+msListaDeCampos[i]);
      msIBOTabela.SQL.Add(')');

      msIBOTabela.SQL.Add('values (:'+msListaDeCampos[0]);
      for i := 1 to msListaDeCampos.Count -1 do
        msIBOTabela.SQL.Add(', :'+msListaDeCampos[i]);
      msIBOTabela.SQL.Add(')');
    end;
    uiUpdate:
    begin
      msIBOTabela.SQL.Add('Update '+msTabela+' set '+msListaDeCampos[0]+' = :'+msListaDeCampos[0]);
      for i := 1 to msListaDeCampos.Count -1 do
        msIBOTabela.SQL.Add(', '+msListaDeCampos[i]+' = :'+msListaDeCampos[i]);
    end;
  end;
end;

procedure TfDB.FormataCampo(fcTipoDoCampo:word; fcCampo:TField; fcDisplayFormat:String; fcAlinhamento:TAlignment);
begin
  case fcTipoDoCampo of
    TipoTimeStamp : begin
                      TDateTimeField(fcCampo).DisplayFormat := fcDisplayFormat;
                      TDateTimeField(fcCampo).Alignment     := fcAlinhamento;
                    end;
    TipoDate      : begin
                      TDateField(fcCampo).DisplayFormat := fcDisplayFormat;
                      TDateField(fcCampo).Alignment     := fcAlinhamento
                    end;
    TipoNumeric   : begin
                      TCurrencyField(fcCampo).DisplayFormat := fcDisplayFormat;
                      TCurrencyField(fcCampo).Alignment     := fcAlinhamento
                    end;
    TipoInteger   : begin
                      TIntegerField(fcCampo).DisplayFormat := fcDisplayFormat;
                      TIntegerField(fcCampo).Alignment     := fcAlinhamento
                    end;
    TipoSmallint  : begin
                      TSmallintField(fcCampo).DisplayFormat := fcDisplayFormat;
                      TSmallintField(fcCampo).Alignment     := fcAlinhamento
                    end;
    TipoVarChar   : begin
                      TStringField(fcCampo).EditMask  := fcDisplayFormat;
                      TStringField(fcCampo).Alignment := fcAlinhamento
                    end;
    TipoChar      : begin
                      TStringField(fcCampo).EditMask  := fcDisplayFormat;
                      TStringField(fcCampo).Alignment := fcAlinhamento
                    end;
  end;
end;

function TfDB.ExisteParametro(epTabela:TIB_Query; epParametro:String):Boolean;
var
  i : integer;
  Tem : boolean;
begin
  Tem := false;
  for i := 0 to epTabela.ParamCount-1 do
    if uppercase(epTabela.Params[i].FieldName) = uppercase(epParametro) then
    begin
      Tem := true;
      break;
    end;
  result := tem;
end;

function TfDB.ExisteParametro(epTabela:TIBOQuery; epParametro:String):Boolean;
var
  i : integer;
  Tem : boolean;
begin
  Tem := false;
  for i := 0 to epTabela.ParamCount-1 do
    if uppercase(epTabela.Params[i].Name) = uppercase(epParametro) then
    begin
      Tem := true;
      break;
    end;
  result := tem;
end;

procedure TfDB.PegaCampos(pcNomeDaTabela: String; var Lista:TStrings; pcConexao:TConexao=tcServer);
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, pcConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('SELECT distinct RDB$RELATION_FIELDS.RDB$FIELD_NAME AS CAMPO');
    EstaTabela.SQL.Add('FROM RDB$RELATION_FIELDS, RDB$FIELDS');
    EstaTabela.SQL.Add('WHERE ( RDB$RELATION_FIELDS.RDB$RELATION_NAME = :T )');
    EstaTabela.ParamByName('T').AsString := pcNomeDaTabela;
    EstaTabela.Active := true;
    Lista.Clear;
    while not EstaTabela.eof do
    begin
      Lista.Add(EstaTabela.FieldByName('CAMPO').AsString);
      EstaTabela.next;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.TemCampoNumTQuery(tcTabela:TIBOQuery; tcNomeDoCampo:String):boolean;
var
  s : string;
begin
  result := true;
  try
    s := tcTabela.FieldByName(tcNomeDoCampo).AsString;
  except
    result := false;
  end;
end;

function TfDB.ExisteTabela(ttNomeDaTabela:String; tTConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, tTConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('SELECT DISTINCT RDB$RELATION_FIELDS.RDB$RELATION_NAME AS TABELAS');
    EstaTabela.SQL.Add('FROM RDB$RELATION_FIELDS, RDB$FIELDS');
    EstaTabela.SQL.Add('WHERE ( RDB$RELATION_FIELDS.RDB$FIELD_SOURCE = RDB$FIELDS.RDB$FIELD_NAME )');
    EstaTabela.SQL.Add(' AND ( RDB$FIELDS.RDB$SYSTEM_FLAG <> 1 )');
    EstaTabela.SQL.Add(' AND (RDB$RELATION_FIELDS.RDB$RELATION_NAME = :N)');
    EstaTabela.ParamByName('N').AsString := ttNomeDaTabela;
    EstaTabela.Active := true;
    result :=  not EstaTabela.IsEmpty;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.ExisteCampo(tcNomeDoCampo, tcNomeDaTabela:String; tcTipoConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, tcTipoConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('SELECT distinct RDB$RELATION_FIELDS.RDB$FIELD_NAME AS CAMPO');
    EstaTabela.SQL.Add('FROM RDB$RELATION_FIELDS, RDB$FIELDS');
    EstaTabela.SQL.Add('WHERE ( RDB$RELATION_FIELDS.RDB$RELATION_NAME = :T )');
    EstaTabela.SQL.Add('  AND ( RDB$RELATION_FIELDS.RDB$FIELD_NAME = :N )');
    EstaTabela.ParamByName('T').AsString := tcNomeDaTabela;
    EstaTabela.ParamByName('N').AsString := tcNomeDoCampo;
    EstaTabela.Active := true;
    result :=  not EstaTabela.IsEmpty;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

procedure TfDB.EsvaziaTabela(etTabela:String; etCondicao:String; eTConexao:TConexao=tcServer);
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, eTConexao);

    if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
    try
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('Delete from '+ etTabela);
      if not Vazio(etCondicao) then EstaTabela.SQL.Add('where '+ etCondicao);
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ExecSQL;
      EstaTabela.IB_Transaction.Commit;
    except
      on E: exception do begin
        CM.LogDeErros(E.Message);
        UltimoErro := E.Message;
        EstaTabela.IB_Transaction.Rollback;
        CM.MensagemDeErro('Não foi possível esvaziar esta tabela');
      end;
    end;
  finally
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.TabelaTaVazia(etTabela:String; etCondicao:String; eTConexao:TConexao=tcServer):Boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, eTConexao);

    if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
    try
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.Add('Select count(*) as Q from '+ etTabela);
      if not Vazio(etCondicao) then EstaTabela.SQL.Add('where '+ etCondicao);
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.Active := true;
      Result := EstaTabela.FieldByName('Q').AsInteger = 0;
      EstaTabela.Active := true;
    except
      on E: exception do begin
        result := True;
        CM.LogDeErros(E.Message);
        UltimoErro := E.Message;
        EstaTabela.IB_Transaction.Rollback;
        CM.MensagemDeErro('Não foi possível esvaziar esta tabela');
      end;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.TamanhoDoCampo(tcTabela, tcCampo: String; tcTipoConexao:TConexao=tcServer): integer;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, tcTipoConexao);

    with EstaTabela do begin
      if not IB_Transaction.InTransaction then IB_Transaction.StartTransaction;
      Active := false;
      SQL.Clear;
      SQL.Add('select rdb$field_length from rdb$fields');
      SQL.Add('where rdb$fields.RDB$FIELD_NAME =');
      SQL.Add('   (select RDB$FIELD_SOURCE from rdb$relation_fields');
      SQL.Add('    where rdb$relation_fields.rdb$RELATION_NAME = '+QuotedStr(tcTabela));
      SQL.Add('    and   rdb$relation_fields.rdb$FIELD_NAME = '+QuotedStr(tcCampo)+')');
      Active := true;
      if not IsEmpty then result := FieldByName('rdb$field_length ').AsInteger
                     else result := CodigoVazio;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.TabelaDaQuery(tqQuery:TStrings):String;
var
  i : integer;
  Texto : String;
begin
  result := '';
  Texto := '';
  for i := 0 to tqQuery.Count -1 do
    Texto := Texto + uppercase(tqQuery[i])+ ' ';

  Texto := SubstituiChar(Texto, ['(', ')'], ' ', true);
  i := pos(' FROM ', Texto);
  Delete(Texto, 1, i+5);

  result := PrimeiraPalavra(Texto);
end;

procedure TfDB.CriaRegistroSeNaoExistir(crTabela, crCampo:String; crCampoValor:Variant; crCampos:Array of String; crValores:Array of Variant; crConexao:TConexao=tcServer);
var
  NaoExiste : Boolean;
  i : integer;
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, crConexao);

    crCampo := TiraEspacos(crCampo);
    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('Select '+crCampo+' from '+crTabela);
    if crCampo <> '*' then
    begin
      EstaTabela.SQL.Add('where '+crCampo+' = :'+crCampo);
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.ParamByName(crCampo).AsVariant := crCampoValor;
    end;
    if not EstaTabela.Prepared then EstaTabela.Prepare;
    EstaTabela.Active := true;
    NaoExiste := EstaTabela.IsEmpty;
    EstaTabela.IB_Transaction.Commit;
    EstaTabela.Active := false;

    if NaoExiste then
    begin
      if not EstaTabela.IB_Transaction.InTransaction then EstaTabela.IB_Transaction.StartTransaction;
      try
        EstaTabela.Active := false;
        EstaTabela.SQL.Clear;
        EstaTabela.SQL.Add('Insert into '+crTabela+' ('+crCampos[0]);
        for i := 1 to Length(crCampos)-1 do
          EstaTabela.SQL.Add(', '+crCampos[i]);
        EstaTabela.SQL.Add(') Values (:'+crCampos[0]);
        for i := 1 to Length(crCampos)-1 do
          EstaTabela.SQL.Add(', :'+crCampos[i]);
        EstaTabela.SQL.Add(')');
        if not EstaTabela.Prepared then EstaTabela.Prepare;

        for i := 0 to Length(crCampos)-1 do
          EstaTabela.ParamByName(crCampos[i]).AsVariant := crValores[i];

        EstaTabela.ExecSQL;
        EstaTabela.IB_Transaction.Commit;
      except
        on E: exception do begin
          UltimoErro := E.Message;
          CM.LogDeErros(E.Message);
          EstaTabela.IB_Transaction.RollBack;
        end;
      end;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.CampoEChavePrimaria(pdTabela,pdCampo:String; pdConexao:TConexao=tcServer):boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
  Chave: String;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, pdConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('select S.RDB$INDEX_NAME CHAVE from RDB$INDEX_SEGMENTS S');
    EstaTabela.SQL.Add('join RDB$INDICES I on I.RDB$INDEX_NAME = S.RDB$INDEX_NAME');
    EstaTabela.SQL.Add('where I.RDB$RELATION_NAME = :TABELA');
    EstaTabela.SQL.Add('and S.RDB$FIELD_NAME = :CAMPO');
    EstaTabela.SQL.Add('and I.RDB$UNIQUE_FLAG = 1');
    if not EstaTabela.Prepared then EstaTabela.Prepare;
    EstaTabela.ParamByName('TABELA').AsString := pdTabela;
    EstaTabela.ParamByName('CAMPO').AsString := pdCampo;
    EstaTabela.Active := true;
    if not EstaTabela.IsEmpty then
      Chave := EstaTabela.FieldByName('CHAVE').AsString;

    result := not Vazio(Chave);
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.PegaCamposChavePrimaria(pdTabela:String; pdConexao:TConexao=tcServer):TResultVetor;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    setLength(result, 0);
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, pdConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('select RDB$FIELD_NAME CAMPO from RDB$INDEX_SEGMENTS S');
    EstaTabela.SQL.Add('join RDB$INDICES I on I.RDB$INDEX_NAME = S.RDB$INDEX_NAME');
    EstaTabela.SQL.Add('where I.RDB$RELATION_NAME = :TABELA');
    EstaTabela.SQL.Add('and I.RDB$UNIQUE_FLAG = 1');
    if not EstaTabela.Prepared then EstaTabela.Prepare;
    EstaTabela.ParamByName('TABELA').AsString := pdTabela;
    EstaTabela.Active := true;

    while not EstaTabela.eof do
    begin
      setLength(result, length(result)+1);
      result[length(result)-1] := EstaTabela.FieldByName('CAMPO').AsString;
      EstaTabela.Next;
    end;

  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

procedure TfDB.PegaTextoDoCampo(var ptTexto:TStrings; ptCampo, ptTabela, ptCondicao: String; ptConexao:TConexao=tcServer);
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, ptConexao);

    EstaTabela.Active := false;
    EstaTabela.SQL.Clear;
    EstaTabela.SQL.Add('Select '+ptCampo+' as RESULTADO from '+ptTabela);
    if not vazio(ptCondicao) then
      EstaTabela.SQL.Add('where '+ptCondicao);

    EstaTabela.Active := true;

    EstaTabela.FieldByName('RESULTADO').AssignTo(ptTexto);
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.TipoDoCampo(tcTabela, tcCampo: String; tcTipoConexao:TConexao=tcServer): integer;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  if ExistePalavra(tcCampo, '.', true) then tcCampo := PegaAPartirDoUltimo(tcCampo, '.');
  try
    fDB.CreateQueryTransacao(EstaTabela, EstaTransacao, tcTipoConexao);

    with EstaTabela do begin
      if not IB_Transaction.InTransaction then IB_Transaction.StartTransaction;
      Active := false;
      SQL.Clear;
      SQL.Add('select RDB$FIELD_TYPE from rdb$fields');
      SQL.Add('where rdb$fields.RDB$FIELD_NAME =');
      SQL.Add('   (select RDB$FIELD_SOURCE from rdb$relation_fields');
      SQL.Add('    where rdb$relation_fields.rdb$RELATION_NAME = '+QuotedStr(tcTabela));
      SQL.Add('    and   rdb$relation_fields.rdb$FIELD_NAME = '+QuotedStr(tcCampo)+')');
      Active := true;
      if not IsEmpty then result := FieldByName('RDB$FIELD_TYPE').AsInteger
                     else result := CodigoVazio;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    fDB.FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

function TfDB.CriaCampo(ccConexao:TConexao; ccTabela:string; ccCampo:string; ccDomain:String):Boolean;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  try
    try
      CreateQueryTransacao(EstaTabela, EstaTransacao, ccConexao);
      EstaTabela.IB_Transaction.StartTransaction;
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.add('alter table '+ccTabela+' add '+ccCampo+' '+ccDomain);
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.Active := true;
      EstaTabela.IB_Transaction.Commit;
      Result := true;
    except
      Result := False;
      EstaTabela.IB_Transaction.Rollback;
    end;
  finally
    fDB.FechaTT(EstaTabela);
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;


function TfDB.PegaVersaoDoFirebird(pvConexao:TConexao=tcServer):String;
var
  EstaTabela : TIB_Query;
  EstaTransacao : TIB_Transaction;
begin
  result := '';
  try
    CreateQueryTransacao(EstaTabela, EstaTransacao, pvConexao);
    try
      EstaTabela.IB_Transaction.StartTransaction;
      EstaTabela.Active := false;
      EstaTabela.SQL.Clear;
      EstaTabela.SQL.add('SELECT RDB$GET_CONTEXT(''SYSTEM'', ''ENGINE_VERSION'') AS VERSAO FROM RDB$DATABASE');
      if not EstaTabela.Prepared then EstaTabela.Prepare;
      EstaTabela.Active := true;
      result := EstaTabela.FieldByName('VERSAO').AsString;
    except
      result := '';
    end;
  finally
    fDB.FechaTT(EstaTabela);
    FreeQueryTransacao(EstaTabela, EstaTransacao);
  end;
end;

procedure TfDB.OrdenaDataSetGrid(var CDS:TClientDataSet; Column:TColumn; var dbgPrin:TDBGrid);
const
  idxDefault = 'DEFAULT_ORDER';
var
  strColumn : string;
  i : integer;
  bolUsed : boolean;
  idOptions : TIndexOptions;
begin

  strColumn := idxDefault;

  if Column.Field.FieldKind in [fkCalculated, fkLookup, fkAggregate] then
    Exit;

  if Column.Field.DataType in [ftBlob, ftMemo] then Exit;

  for i := 0 to dbgPrin.Columns.Count -1 do
    dbgPrin.Columns[i].Title.Font.Style := [];

  bolUsed := (Column.Field.FieldName = CDS.IndexName);

  CDS.IndexDefs.Update;
  for i := 0 to CDS.IndexDefs.Count - 1 do
  begin
    if CDS.IndexDefs.Items[i].Name = Column.Field.FieldName then
    begin
      strColumn := Column.Field.FieldName;
      case (CDS.IndexDefs.Items[i].Options = [ixDescending]) of
        true : idOptions := [];
        false : idOptions := [ixDescending];
      end;
    end;
  end;

  if (strColumn = idxDefault)  or (bolUsed) then
  begin
    if bolUsed then
      CDS.DeleteIndex(Column.Field.FieldName);
    try
      CDS.AddIndex(Column.Field.FieldName, Column.Field.FieldName, idOptions, '', '', 0);
      strColumn := Column.Field.FieldName;
    except
      if bolUsed then
      strColumn := idxDefault;
    end;
  end;

  try
   CDS.IndexName := strColumn;
   Column.Title.Font.Style := [fsbold];
  except
   CDS.IndexName := idxDefault;
  end;
end;

procedure TfDB.GridZebrado (RecNo:LongInt; Grid:TDBGrid; Rect:TRect;
  State:TGridDrawState; Column:TColumn; Colunaordenada:TColumn=nil;
  CorSim:TColor=clMoneyGreen; CorNao:TColor=clWhite; CorSelecionado:TColor=$00619FE4);
begin
  if not(Odd(RecNo)) then // Se for Ìmpar
  begin
    with Grid do
    begin
      with Canvas do
      begin
        if Column = ColunaOrdenada then
          Brush.Color := IntensificaCor(CorSim, -30)
        else
          Brush.Color := CorSim;

        FillRect (Rect); // Pinta a célula
      end;
      DefaultDrawDataCell (Rect, Column.Field, State) // Reescreve o valor que vem do banco
    end;
  end
  else begin              // Se for Par
    with Grid do
    begin
      with Canvas do
      begin
        if Column = ColunaOrdenada then
          Brush.Color := IntensificaCor(CorNao, -30)
        else
          Brush.Color := CorNao;

        FillRect (Rect); // Pinta a célula
      end;
      DefaultDrawDataCell (Rect, Column.Field, State) // Reescreve o valor que vem do banco
    end;
  end;

  if (gdSelected in State) then  // Se a Linha ou a célula estiver Selecionada
  begin
    with Grid do
    begin
      with Canvas do
      begin
        Brush.Color := CorSelecionado;
        Font.Color := clBlack;
        FillRect (Rect); // Pinta a célula
      end;
      DefaultDrawDataCell (Rect, Column.Field, State) // Reescreve o valor que vem do banco
    end;
  end;
end;

function TfDB.SomaColunaCDS(var CDS:TClientDataSet; Coluna:String):Real;
var
  Book : TBookmark;
begin
  try
    Book := CDS.GetBookmark;
    CDS.DisableControls;
    CDS.First;
    result := 0;
    while not CDS.eof do
    begin
      result := result + CDS.FieldByName(Coluna).AsFloat;
      CDS.Next;
    end;
  finally
    CDS.GotoBookmark(Book);
    CDS.EnableControls;
  end;
end;

procedure TfDB.DestacaColunaOrdenada(Grid:TDBGrid; Rect:TRect; State:TGridDrawState; DataCol:Integer; Column:TColumn; ColunaOrdenada:TColumn);
begin
  if Column = ColunaOrdenada then
  begin
    Grid.Canvas.Brush.Color := IntensificaCor(Grid.Canvas.Brush.Color, -30);
    Grid.DefaultDrawColumnCell(Rect, DataCol, Column, State);
  end;
end;


end.
