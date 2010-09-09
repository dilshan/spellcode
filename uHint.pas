//------------------------------------------------------------------
// SpellCode - Source code spell checking utility. 
// Copyright © 2010 Dilshan R Jayakody. <jayakody2000lk@live.com>
// http://code.google.com/p/spellcode
// 
// SpellCode is distributed under the MIT license.
//------------------------------------------------------------------

unit uHint;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons;

type
  TfmHint = class(TForm)
    shpMain: TShape;
    btnClose: TSpeedButton;
    lblWord: TLabel;
    lstSugs: TListBox;
    btnAdd: TSpeedButton;
    btnIgnore: TSpeedButton;
    btnOK: TSpeedButton;
    btnIgnoreAll: TSpeedButton;
    procedure btnIgnoreClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure btnIgnoreAllClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmHint: TfmHint;

implementation

{$R *.dfm}

procedure TfmHint.btnIgnoreClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TfmHint.btnOKClick(Sender: TObject);
begin
  ModalResult := mrOK;
end;

procedure TfmHint.btnCloseClick(Sender: TObject);
begin
  ModalResult := mrNo;
end;

procedure TfmHint.btnAddClick(Sender: TObject);
begin
  ModalResult := mrYes;
end;

procedure TfmHint.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if(Key = 119) then
    ModalResult := mrCancel
  else if(Key = 13) then
    ModalResult := mrOK
  else if(Key = 27) then
    ModalResult := mrNo
  else if(Key = 45) then
    ModalResult := mrYes
  else if(Key = 120) then
    ModalResult := mrYesToAll
end;

procedure TfmHint.FormShow(Sender: TObject);
begin
  btnOK.Enabled := (lstSugs.Items.Count > 0);
  lstSugs.Enabled := btnOK.Enabled;
end;

procedure TfmHint.btnIgnoreAllClick(Sender: TObject);
begin
  ModalResult := mrYesToAll;
end;

end.
