unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Mask, sMaskEdit, sCustomComboEdit, sTooledit, StdCtrls, DB,
  IBCustomDataSet, IBQuery, IBDatabase, frxDesgn,
  frxClass, frxExportXLS, frxDBSet, frxExportImage, Grids, DBGrids;

type
  TForm2 = class(TForm)
    Terminal: TIBDatabase;
    IBTransaction1: TIBTransaction;
    TQuery1: TIBQuery;
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    DateFrom: TsDateEdit;
    DateTo: TsDateEdit;
    frxReport1: TfrxReport;
    frxDBDataset1: TfrxDBDataset;
    frxXLSExport1: TfrxXLSExport;
    frxDesigner1: TfrxDesigner;
    sDirectoryEdit1: TsDirectoryEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  Path: string;

implementation

uses Unit1;

{$R *.dfm}

procedure TForm2.Button1Click(Sender: TObject);
begin
Memo1.Lines.Add('Выгружаем кассы ОПЕРУ...');
  with TQuery1 do
  Begin
    SQL.Clear;
    SQL.Text:='select terminal.name, smena.id, smena.data_n, smena.data_k, count(oplata.summa), sum(oplata.summa)'
    +' from oplata'
    +' inner join smena on oplata.kod_smen=smena.id'
    +' inner join terminal on oplata.id_terminal=terminal.id'
    +' where terminal.bank_id=4 and terminal.rs_debet like ''1001%'' and cast(oplata.data_v as Date)'
    +' between '''+DateToStr(DateFrom.Date)+''' and  '''+DateToStr(DateTo.Date)+''' and oplata.storno_id=0'
    +' group by terminal.name, smena.id, smena.data_n, smena.data_k';
    Open;
End;
frxReport1.LoadFromFile(Path+'Kassa.fr3');
frxReport1.Variables.Clear;
frxReport1.Variables['Date_from']:=DateFrom.Date;
frxReport1.Variables['Date_To']:=DateTo.Date;
frxReport1.DesignReport();
//frxReport1.PrepareReport();
if DateFrom.Date <> DateTo.Date
  then
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Реестр по кассам (ОПЕРУ) за период с '+DateToStr(DateFrom.Date)+' по ' +DateToStr(DateTo.Date)+'.xls'
  else
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Реестр по кассам (ОПЕРУ) за '+DateToStr(DateTo.Date)+'.xls';
frxReport1.Export(frxXLSExport1);
Memo1.Lines.Add('Файл '+frxXLSExport1.FileName+' выгружен.');
Memo1.Lines.Add('Выгружаем кассы (город)...');
with TQuery1 do
  Begin
    SQL.Clear;
    SQL.Text:='select terminal.name, smena.id, smena.data_n, smena.data_k, count(oplata.summa), sum(oplata.summa)'
    +' from oplata'
    +' inner join smena on oplata.kod_smen=smena.id'
    +' inner join terminal on oplata.id_terminal=terminal.id'
    +' where terminal.bank_id=4 and terminal.rs_debet not like ''1001%'' and terminal.rs_credit=2902690003 and cast(oplata.data_v as Date)'
    +' between '''+DateToStr(DateFrom.Date)+''' and  '''+DateToStr(DateTo.Date)+''' and oplata.storno_id=0'
    +' group by terminal.name, smena.id, smena.data_n, smena.data_k';
    Open;
End;
frxReport1.LoadFromFile(Path+'Kassa.fr3');
frxReport1.Variables.Clear;
frxReport1.Variables['Date_from']:=DateFrom.Date;
frxReport1.Variables['Date_To']:=DateTo.Date;
//frxReport1.DesignReport();
frxReport1.PrepareReport();
if DateFrom.Date <> DateTo.Date
  then
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Реестр по кассам (город) за период с '+DateToStr(DateFrom.Date)+' по ' +DateToStr(DateTo.Date)+'.xls'
  else
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Реестр по кассам (город) за '+DateToStr(DateTo.Date)+'.xls';
frxReport1.Export(frxXLSExport1);
Memo1.Lines.Add('Файл '+frxXLSExport1.FileName+' выгружен.');
Memo1.Lines.Add('Выгружаем кассы...');
with TQuery1 do
  Begin
    SQL.Clear;
    SQL.Text:='select terminal.name, smena.id, smena.data_n, smena.data_k, count(oplata.summa), sum(oplata.summa)'
    +' from oplata'
    +' inner join smena on oplata.kod_smen=smena.id'
    +' inner join terminal on oplata.id_terminal=terminal.id'
    +' where terminal.bank_id=4 and terminal.rs_debet not like ''1001%'' and terminal.rs_credit<>2902690003 and cast(oplata.data_v as Date)'
    +' between '''+DateToStr(DateFrom.Date)+''' and  '''+DateToStr(DateTo.Date)+''' and oplata.storno_id=0'
    +' group by terminal.name, smena.id, smena.data_n, smena.data_k';
    Open;
End;
frxReport1.LoadFromFile(Path+'Kassa.fr3');
frxReport1.Variables.Clear;
frxReport1.Variables['Date_from']:=DateFrom.Date;
frxReport1.Variables['Date_To']:=DateTo.Date;
//frxReport1.DesignReport();
frxReport1.PrepareReport();
if DateFrom.Date <> DateTo.Date
  then
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Реестр по кассам за период с '+DateToStr(DateFrom.Date)+' по ' +DateToStr(DateTo.Date)+'.xls'
  else
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Реестр по кассам за '+DateToStr(DateTo.Date)+'.xls';
frxReport1.Export(frxXLSExport1);
Memo1.Lines.Add('Файл '+frxXLSExport1.FileName+' выгружен.');
end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TForm2.Button2Click(Sender: TObject);
begin
Memo1.Lines.Add('Выгружаем сводный по организациям ОПЕРУ...');
with TQuery1 do
  Begin
    SQL.Clear;
    SQL.Text:='select oplata.kod_org, organization.name,  sum(oplata.summa), count(oplata.summa),'
    +' komissia.percent_our, sum(oplata.komisiya)/100, komissia.percent_out, sum(oplata.komisiya_bank)/100,'
    +' sum(oplata.summa)-sum(oplata.komisiya)/100-sum(oplata.komisiya_bank)/100'
    +' from oplata'
    +' inner join komissia on oplata.kod_org=komissia.kod_org'
    +' inner join organization on oplata.kod_org=organization.id'
    +' inner join terminal on oplata.id_terminal=terminal.id'
    +' where terminal.bank_id=4 and komissia.bank_id=4 and terminal.rs_debet like ''1001%'' and cast(oplata.data_v as Date)'
    +' between '''+DateToStr(DateFrom.Date)+''' and '''+DateToStr(DateTo.Date)+''' and oplata.usluga=komissia.usluga and oplata.usluga<>-1 and oplata.storno_id=0'
    +' group by oplata.kod_org, organization.name, komissia.percent_our, komissia.percent_out';
    Open;
End;

frxReport1.LoadFromFile(Path+'Org.fr3');
frxReport1.Variables.Clear;
frxReport1.Variables['Date_from']:=DateFrom.Date;
frxReport1.Variables['Date_To']:=DateTo.Date;
//frxReport1.DesignReport();
frxReport1.PrepareReport();
if DateFrom.Date <> DateTo.Date
  then
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Сводный по организациям (ОПЕРУ) за период с '+DateToStr(DateFrom.Date)+' по ' +DateToStr(DateTo.Date)+'.xls'
  else
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Сводный по организациям (ОПЕРУ) за '+DateToStr(DateTo.Date)+'.xls';
frxReport1.Export(frxXLSExport1);
Memo1.Lines.Add('Файл '+frxXLSExport1.FileName+' выгружен.');
////////////////////////////////////////////////////////////////////////////////
Memo1.Lines.Add('Выгружаем сводный по организациям(город)...');
with TQuery1 do
  Begin
    SQL.Clear;
    SQL.Text:='select oplata.kod_org, organization.name,  sum(oplata.summa), count(oplata.summa),'
    +' komissia.percent_our, sum(oplata.komisiya)/100, komissia.percent_out, sum(oplata.komisiya_bank)/100,'
    +' sum(oplata.summa)-sum(oplata.komisiya)/100-sum(oplata.komisiya_bank)/100'
    +' from oplata'
    +' inner join komissia on oplata.kod_org=komissia.kod_org'
    +' inner join organization on oplata.kod_org=organization.id'
    +' inner join terminal on oplata.id_terminal=terminal.id'
    +' where terminal.bank_id=4 and komissia.bank_id=4 and terminal.rs_debet not like ''1001%'' and terminal.rs_credit=2902690003 and cast(oplata.data_v as Date)'
    +' between '''+DateToStr(DateFrom.Date)+''' and '''+DateToStr(DateTo.Date)+''' and oplata.usluga=komissia.usluga and oplata.usluga<>-1 and oplata.storno_id=0'
    +' group by oplata.kod_org, organization.name, komissia.percent_our, komissia.percent_out';
    Open;
End;

frxReport1.LoadFromFile(Path+'Org.fr3');
frxReport1.Variables.Clear;
frxReport1.Variables['Date_from']:=DateFrom.Date;
frxReport1.Variables['Date_To']:=DateTo.Date;
//frxReport1.DesignReport();
frxReport1.PrepareReport();
if DateFrom.Date <> DateTo.Date
  then
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Сводный по организациям (город) за период с '+DateToStr(DateFrom.Date)+' по ' +DateToStr(DateTo.Date)+'.xls'
  else
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Сводный по организациям (город) за '+DateToStr(DateTo.Date)+'.xls';
frxReport1.Export(frxXLSExport1);
Memo1.Lines.Add('Файл '+frxXLSExport1.FileName+' выгружен.');
////////////////////////////////////////////////////////////////////////////////
Memo1.Lines.Add('Выгружаем сводный по организациям...');
with TQuery1 do
  Begin
    SQL.Clear;
    SQL.Text:='select oplata.kod_org, organization.name,  sum(oplata.summa), count(oplata.summa),'
    +' komissia.percent_our, sum(oplata.komisiya)/100, komissia.percent_out, sum(oplata.komisiya_bank)/100,'
    +' sum(oplata.summa)-sum(oplata.komisiya)/100-sum(oplata.komisiya_bank)/100'
    +' from oplata'
    +' inner join komissia on oplata.kod_org=komissia.kod_org'
    +' inner join organization on oplata.kod_org=organization.id'
    +' inner join terminal on oplata.id_terminal=terminal.id'
    +' where terminal.bank_id=4 and komissia.bank_id=4 and terminal.rs_debet not like ''1001%'' and terminal.rs_credit<>2902690003 and cast(oplata.data_v as Date)'
    +' between '''+DateToStr(DateFrom.Date)+''' and '''+DateToStr(DateTo.Date)+''' and oplata.usluga=komissia.usluga and oplata.usluga<>-1 and oplata.storno_id=0'
    +' group by oplata.kod_org, organization.name, komissia.percent_our, komissia.percent_out';
    Open;
End;

frxReport1.LoadFromFile(Path+'Org.fr3');
frxReport1.Variables.Clear;
frxReport1.Variables['Date_from']:=DateFrom.Date;
frxReport1.Variables['Date_To']:=DateTo.Date;
//frxReport1.DesignReport();
frxReport1.PrepareReport();
if DateFrom.Date <> DateTo.Date
  then
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Сводный по организациям за период с '+DateToStr(DateFrom.Date)+' по ' +DateToStr(DateTo.Date)+'.xls'
  else
    frxXLSExport1.FileName:=sDirectoryEdit1.Text+'\'+'Сводный по организациям за '+DateToStr(DateTo.Date)+'.xls';
frxReport1.Export(frxXLSExport1);
Memo1.Lines.Add('Файл '+frxXLSExport1.FileName+' выгружен.');
end;

procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Form1.Close;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
Terminal.DatabaseName:=Form1.ComboBox1.Text+':terminal';
 try
    Terminal.Connected := True;

  except
    on e: exception do
      Application.MessageBox('Ошибка подключения', 'Ошибка', MB_OK + MB_iconerror);
  end;
  Path:= ExtractFilePath(Application.ExeName);
  sDirectoryEdit1.Text:=Path+'Reports';
  if DirectoryExists(sDirectoryEdit1.Text)<>True then
        begin
          ForceDirectories(sDirectoryEdit1.Text);
        end;
DateFrom.Date:=date-1;
DateTo.Date:=date-1;
Memo1.Lines.Clear;
end;

end.
