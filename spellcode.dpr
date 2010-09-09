//------------------------------------------------------------------
// SpellCode - Source code spell checking utility. 
// Copyright © 2010 Dilshan R Jayakody. <jayakody2000lk@live.com>
// http://code.google.com/p/spellcode
// 
// SpellCode is distributed under the MIT license.
//------------------------------------------------------------------

program spellcode;

uses
  Forms,
  uMain in 'uMain.pas' {fmMain},
  uHint in 'uHint.pas' {fmHint},
  uAbout in 'uAbout.pas' {fmAbout},
  uStartup in 'uStartup.pas' {uStartupWin};

{$E exe}

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'SpellCode';
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
