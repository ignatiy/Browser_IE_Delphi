unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.OleCtrls, SHDocVw,
  Vcl.Menus, Vcl.ExtCtrls, Vcl.Buttons, Vcl.ComCtrls, MSHTML, ActiveX, Vcl.Imaging.pngimage;

type
  TForm1 = class(TForm)
    ComboBox1: TComboBox;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    Timer1: TTimer;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    ProgressBar1: TProgressBar;
    Image1: TImage;
    Image2: TImage;
    Image3: TImage;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    N4: TMenuItem;
    OpenDialog1: TOpenDialog;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    SaveDialog1: TSaveDialog;
    Timer2: TTimer;
    Label1: TLabel;
    SpeedButton1: TSpeedButton;
    PageControl1: TPageControl;
    Panel1: TPanel;
    WebBrowser1: TWebBrowser;
    procedure ComboBox1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure N3Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure WebBrowser1ProgressChange(ASender: TObject; Progress,
      ProgressMax: Integer);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure WebBrowser1NavigateComplete2(ASender: TObject;
      const pDisp: IDispatch; const URL: OleVariant);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure WebBrowser1DocumentComplete(ASender: TObject;
      const pDisp: IDispatch; const URL: OleVariant);
    procedure N10Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);



  private
    { Private declarations }
  public
    { Public declarations }

    procedure SaveHTMLSourceToFile(const FileName: string; WB: TWebBrowser);

  end;

var
  Form1: TForm1;
  x,y:integer;
  web: array of TWebBrowser;
  tb : array of TTabSheet;
  p: tpagecontrol;


implementation

{$R *.dfm}

uses Unit2 , ShellApi;


function ExecuteScript(doc: IHTMLDocument2; script: string; language: string): Boolean;
var
  win: IHTMLWindow2;
  Olelanguage: Olevariant;
begin
  if doc <> nil then
  begin
    try
      win := doc.parentWindow;
      if win <> nil then
      begin
        try
          Olelanguage := language;
          win.ExecScript(script, Olelanguage);
        finally
          win := nil;
        end;
      end;
    finally
      doc := nil;
    end;
  end;
end;


procedure TForm1.BitBtn1Click(Sender: TObject);
begin
WebBrowser1.GoBack;
end;

procedure TForm1.SaveHTMLSourceToFile(const FileName: string; WB: TWebBrowser);
var
 PersistStream: IPersistStreamInit;
 FileStream: TFileStream;
 Stream: IStream;
 SaveResult: HRESULT;
begin
 PersistStream := WB.Document as IPersistStreamInit;
 FileStream := TFileStream.Create(FileName, fmCreate);
try
 Stream := TStreamAdapter.Create(FileStream, soReference) as IStream;
 SaveResult := PersistStream.Save(Stream, True);
if FAILED(SaveResult) then
 MessageBox(Handle, 'Fail to save HTML source', 'Error', 0);
finally
 FileStream.Free;
end;
end;


procedure TForm1.SpeedButton1Click(Sender: TObject);
begin
 setlength(tb, length(tb)+1);
 tb[length(tb)-1]:= TTabSheet.Create(self);
 with tb[length(tb)-1] do begin
  Caption:= inttostr(length(tb)-1);
  //height:= 165; width:= 300;
  Align:= alclient;
  tb[length(tb)-1].PageControl:= p;
 end;
 p.ActivePageIndex:= length(tb)-1;

 setlength(web, length(web)+1);
 web[length(web)-1]:= TWebBrowser.Create(self);
 with web[length(web)-1] do begin
  Align:= alclient;
  web[length(web)-1].Navigate('http://www.yandex.ru')
 end;
 tb[length(tb)-1].InsertControl(web[length(web)-1]);
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
WebBrowser1.GoForward;
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
begin
WebBrowser1.Refresh;
end;

procedure TForm1.BitBtn4Click(Sender: TObject);
begin
WebBrowser1.Stop;
end;

procedure TForm1.BitBtn5Click(Sender: TObject);
var
s:OleVariant;
IDoc1: IHTMLDocument2;
begin
if ComboBox1.Text <> '' then WebBrowser1.Navigate(ComboBox1.Text);
Webbrowser1.Document.QueryInterface(IHTMLDocument2, iDoc1);
s:='javascript:';
ExecuteScript(iDoc1,s,'JavaScript');
//WebBrowser1.ScriptErrorsSuppressed = True;
end;

procedure TForm1.BitBtn6Click(Sender: TObject);
begin
//ShellExecute(handle, 'open', 'http://www.yandex.ru/', nil, nil, SW_SHOW);
WebBrowser1.GoHome;

end;

procedure TForm1.ComboBox1Change(Sender: TObject);
begin
WebBrowser1.Navigate(ComboBox1.Text);
end;

procedure TForm1.ComboBox1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if Key= VK_RETURN then
WebBrowser1.Navigate(ComboBox1.Text);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
x:=ClientWidth;
y:=ClientHeight;
WebBrowser1.Navigate('http://www.yandex.ru');
p:= tpagecontrol.Create(self);
 with p do begin
  p.Align:= alclient;
  parent:= panel1
 end;
 setlength(tb, 0);
 setlength(web, 0);
end;

procedure TForm1.N10Click(Sender: TObject);
begin
if SaveDialog1.Execute then

SaveHTMLSourceToFile(SaveDialog1.FileName, WebBrowser1);
end;

procedure TForm1.N2Click(Sender: TObject);
begin
Form2.ShowModal;
end;

procedure TForm1.N3Click(Sender: TObject);
begin
Close;
end;

procedure TForm1.N4Click(Sender: TObject);
begin
WebBrowser1.ExecWB(OLECMDID_PRINT,OLECMDEXECOPT_PROMPTUSER);
end;

procedure TForm1.N5Click(Sender: TObject);
begin
if OpenDialog1.Execute then
WebBrowser1.Navigate(Opendialog1.FileName);
ComboBox1.Text:=OpenDialog1.FileName;
end;

procedure TForm1.N6Click(Sender: TObject);
begin
WebBrowser1.ExecWB(OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT);
end;

procedure TForm1.N7Click(Sender: TObject);
begin
WebBrowser1.ExecWB(OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT);
end;

procedure TForm1.N8Click(Sender: TObject);
begin
WebBrowser1.ExecWB(OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT);
end;

procedure TForm1.N9Click(Sender: TObject);
begin
WebBrowser1.ExecWB(OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT);
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
if (WebBrowser1.Width <> Form1.Width) then WebBrowser1.Width:=Form1.Width - 27;
if (WebBrowser1.Height <> Form1.Height) then WebBrowser1.Height:=Form1.Height - 135 ;
if (ComboBox1.Width <> Form1.Width) then ComboBox1.Width:=Form1.Width - 100;
if (Progressbar1.Width <> Form1.Width) then Progressbar1.Width:=Form1.Width - 27;
if (Image1.Width <> Form1.Width) then Image1.Width:=Form1.Width - 30;
if (Image2.Width <> Form1.Width) then Image2.Width:=Form1.Width - 30;
if (Image3.Width <> Form1.Width) then Image3.Width:=Form1.Width - 30;
if (BitBtn5.Left <> ComboBox1.Width + 35) then BitBtn5.Left:=ComboBox1.Width + 6;
if (label1.Left <> WebBrowser1.Width + 35) then label1.Left:=WebBrowser1.Width -55;

end;

procedure TForm1.Timer2Timer(Sender: TObject);
begin
label1.Caption := TimeToStr (Time ());
end;

procedure TForm1.WebBrowser1DocumentComplete(ASender: TObject;
  const pDisp: IDispatch; const URL: OleVariant);
begin
if URL=WebBrowser1.LocationURL Then
begin
ShowMessage('Загрузка успешно завершена!');
end;
end;

procedure TForm1.WebBrowser1NavigateComplete2(ASender: TObject;
  const pDisp: IDispatch; const URL: OleVariant);
begin
ComboBox1.Items.Add(WebBrowser1.LocationURL);
end;

procedure TForm1.WebBrowser1ProgressChange(ASender: TObject; Progress,
  ProgressMax: Integer);
begin
Progressbar1.Max:=progressmax;
Progressbar1.Position:=progress;
end;

end.
