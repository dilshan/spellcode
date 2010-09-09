//------------------------------------------------------------------
// SpellCode - Source code spell checking utility. 
// Copyright © 2010 Dilshan R Jayakody. <jayakody2000lk@live.com>
// http://code.google.com/p/spellcode
// 
// SpellCode is distributed under the MIT license.
//------------------------------------------------------------------

unit uStartup;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, JvExControls, JvButton, JvNavigationPane,
  JvXPCore, JvXPBar, StdCtrls, JvExStdCtrls, JvCombobox, JvDriveCtrls,
  JvListBox, ComCtrls, FolderTree, shellapi;

type
  TuStartupWin = class(TForm)
    JvNavPanelHeader2: TJvNavPanelHeader;
    Panel1: TPanel;
    Bevel1: TBevel;
    brLinks: TJvXPBar;
    JvNavPanelHeader1: TJvNavPanelHeader;
    brMRU: TJvXPBar;
    Image1: TImage;
    Label1: TLabel;
    chkStartup: TCheckBox;
    Label2: TLabel;
    cmbLangs: TComboBox;
    btnCancel: TButton;
    btnOK: TButton;
    tvDirs: TFolderTree;
    path_timer1: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure path_timer1Timer(Sender: TObject);
    procedure MRUItemClick(Sender : TObject);
    procedure brLinksItems0Click(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const general_err_txt = 'Critical error has been occurred!'#13;

var
  uStartupWin: TuStartupWin;

implementation
uses uMain;

{$R *.dfm}

procedure TuStartupWin.FormCreate(Sender: TObject);
var strlist : TStringList; pos : integer;
begin
  try
    cmbLangs.Items := fmMain.cmbLanguage.Items;
    cmbLangs.ItemIndex := fmMain.cmbLanguage.ItemIndex;
    Screen.Cursor := crHourGlass;
    strlist := TStringList.Create;
    strlist.LoadFromFile(fmMain.root+'spellcode.mru');
    for pos := 0 to strlist.Count - 1 do begin
      with brMru.Items.Add do begin
        Caption := ExtractFileName(ExcludeTrailingBackslash(strlist.Strings[pos]));
        Hint := strlist.Strings[pos];
        OnClick := MRUItemClick;
      end;
    end;
    freeandnil(strlist);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TuStartupWin.MRUItemClick(Sender : TObject);
begin
  try
    Screen.Cursor := crHourGlass;
    tvDirs.SetDirectory(TJvXPBarItem(Sender).Hint);
    Screen.Cursor := crDefault;
    self.ModalResult := mrOK;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TuStartupWin.path_timer1Timer(Sender: TObject);
begin
  try
    path_timer1.Interval := 0;
    tvDirs.SetDirectory(GetCurrentDir);
    Screen.Cursor := crDefault;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TuStartupWin.brLinksItems0Click(Sender: TObject);
begin
  try
    ShellExecute(Application.Handle,'open',pchar(TJvXPBarItem(Sender).Hint),'','',SW_SHOW);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TuStartupWin.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #27 then begin
    self.ModalResult := mrCancel;
    Key := #0;
  end;
end;

end.
