//------------------------------------------------------------------
// SpellCode - Source code spell checking utility. 
// Copyright © 2010 Dilshan R Jayakody. <jayakody2000lk@live.com>
// http://code.google.com/p/spellcode
// 
// SpellCode is distributed under the MIT license.
//------------------------------------------------------------------

unit uAbout;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, jpeg, shellapi;

type
  TfmAbout = class(TForm)
    Panel1: TPanel;
    btnOK: TButton;
    Bevel1: TBevel;
    Shape1: TShape;
    btnHomePage: TButton;
    Image1: TImage;
    Label1: TLabel;
    lblversion: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    procedure btnHomePageClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmAbout: TfmAbout;

implementation

{$R *.dfm}

procedure TfmAbout.btnHomePageClick(Sender: TObject);
begin
  shellexecute(0,'open','http://code.google.com/p/spellcode','','',SW_SHOW);
end;

procedure TfmAbout.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #27 then begin
    self.ModalResult := mrOK;
    Key := #0;
  end;
end;

end.
