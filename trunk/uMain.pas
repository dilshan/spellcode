//------------------------------------------------------------------
// SpellCode - Source code spell checking utility. 
// Copyright © 2010 Dilshan R Jayakody. <jayakody2000lk@live.com>
// http://code.google.com/p/spellcode
// 
// SpellCode is distributed under the MIT license.
//------------------------------------------------------------------

unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Buttons, StdCtrls, ComCtrls, SynEdit, JvBaseDlg,
  JvBrowseFolder, IniFiles, JclAnsiStrings, SynHighlighterCS, JclShell,
  SynEditHighlighter, SynHighlighterJava, SynHighlighterPHP, uHint, JclFileUtils,
  SynHighlighterXML, SynHighlighterPas, SynHighlighterCpp, registry,
  SynHighlighterRuby, SynHighlighterTclTk, SynHighlighterPython, Clipbrd,
  SynHighlighterVB, SynHighlighterPerl, AspellHeadersDyn, JvExComCtrls,
  JvStatusBar, JvExControls, JvArrowButton, Menus, JclStrings, uStartup,
  JvProgressBar, JvSpecialProgress, treelist, JvExExtCtrls, shellapi,
  JvNetscapeSplitter, uAbout, JclSysInfo, AppEvnts, JclGraphUtils, ImgList,
  JvComponentBase, JvBalloonHint, JvAppStorage, JvAppRegistryStorage,
  JvFormPlacement;

type
  TAppSettings = record
    active_dictionary_id : integer;
    programming_lan_id : integer;
    show_startup_win : Boolean;
    dict_select_message : Boolean;
  end;

  TfmMain = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    btnBrowse: TSpeedButton;
    btnScan: TSpeedButton;
    btnAbout: TSpeedButton;
    btnSettings: TSpeedButton;
    Bevel1: TBevel;
    Label1: TLabel;
    cmbLanguage: TComboBox;
    lstFiles: TTreeView;
    Panel3: TPanel;
    Panel4: TPanel;
    Bevel2: TBevel;
    editor1: TSynEdit;
    Panel5: TPanel;
    lblFilename: TLabel;
    Panel6: TPanel;
    btnSave: TSpeedButton;
    btnCheckSpell: TSpeedButton;
    btnCopy: TSpeedButton;
    btnPaste: TSpeedButton;
    btnCut: TSpeedButton;
    btnUndo: TSpeedButton;
    btnRedo: TSpeedButton;
    dirbrowse: TJvBrowseForFolderDialog;
    hyJava: TSynJavaSyn;
    hyCS: TSynCSSyn;
    hyPHP: TSynPHPSyn;
    hyPython: TSynPythonSyn;
    hyTCL: TSynTclTkSyn;
    hyRuby: TSynRubySyn;
    hyCPP: TSynCppSyn;
    hyPas: TSynPasSyn;
    hyXML: TSynXMLSyn;
    hyPerl: TSynPerlSyn;
    hyVB: TSynVBSyn;
    status_panel: TJvStatusBar;
    btnLanguage: TJvArrowButton;
    popDict: TPopupMenu;
    Bevel3: TBevel;
    btnRescan: TSpeedButton;
    prgLevel: TJvSpecialProgress;
    msgpanel: TPanel;
    msgview1: TTreeList;
    splitter1: TJvNetscapeSplitter;
    msgview_popup: TPopupMenu;
    Ignorealloccurrences1: TMenuItem;
    Addtodictionary1: TMenuItem;
    N1: TMenuItem;
    SaveAs1: TMenuItem;
    Clear1: TMenuItem;
    N2: TMenuItem;
    Locatethemistake1: TMenuItem;
    Getthenumberofoccurrences1: TMenuItem;
    incurrentfile1: TMenuItem;
    inentireproject1: TMenuItem;
    Replacewith1: TMenuItem;
    logsave1: TSaveDialog;
    botbar1: TShape;
    app_evnts: TApplicationEvents;
    fileicon: TImage;
    sugmenu: TPopupMenu;
    glb_images: TImageList;
    N3: TMenuItem;
    Copy1: TMenuItem;
    startup_win: TTimer;
    qhelp: TJvBalloonHint;
    app_frmmain_setup: TJvFormStorage;
    app_reg_settings: TJvAppRegistryStorage;
    function Root : string;
    function LoadAppSettings() : TAppSettings;
    procedure btnBrowseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
	  procedure BuildScanFileList(mainnode : TTreeNode; slist : TStringList; strdir : string; ext_list : TStringList);
	  procedure GetSupportFileExt(file_type : string; ext_list : TStringList);
    procedure lstFilesDblClick(Sender: TObject);
    function GetSyntaxHighlighter(filename : string) : TSynCustomHighlighter;
    procedure LoadASpellDict();
    function GetDictDisplayName(did : string) : string;
    procedure LanguageMenuClick(Sender : TObject);
    procedure SaveAppSettings(settings_ : TAppSettings);
    procedure FormDestroy(Sender: TObject);
    procedure cmbLanguageChange(Sender: TObject);
    procedure btnCheckSpellClick(Sender: TObject);
    function GetActiveDict : string;
    procedure SentzAnalyzer(str : string; wrd_list : TStringList; pos_list : TStringList);
    function is_valid_token(tkid : integer; lan_id : integer) : Boolean;
    procedure btnScanClick(Sender: TObject);
    function GetLanIndex(fname : string) : integer;
    procedure msgview1DblClick(Sender: TObject);
    procedure InterfaceState(state : Boolean);
    function PopBusyToUser() : boolean;
    procedure btnSaveClick(Sender: TObject);
    function CloseEditorSession() : Boolean;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure editor1StatusChange(Sender: TObject;
      Changes: TSynStatusChanges);
    procedure btnAboutClick(Sender: TObject);
    procedure msgview1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure inentireproject1Click(Sender: TObject);
    procedure Addtodictionary1Click(Sender: TObject);
    procedure Clear1Click(Sender: TObject);
    procedure msgview_popupPopup(Sender: TObject);
    procedure replace_word_menu_click(Sender : TObject);
    function is_file_attrib_clr(filename : string; show_all_btns : boolean; var old_status : integer) : Boolean;
    procedure setup_taskbar_button_flash();
    procedure SaveAs1Click(Sender: TObject);
    procedure app_evntsSettingChange(Sender: TObject; Flag: Integer;
      const Section: String; var Result: Integer);
    procedure assign_project_icon(filename : string);
    procedure Getthenumberofoccurrences1Click(Sender: TObject);
    procedure fileiconMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure fileiconDblClick(Sender: TObject);
    procedure sugmenuPopup(Sender: TObject);
    procedure ReplaceWithSpotWord(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure msgview1KeyPress(Sender: TObject; var Key: Char);
    procedure SaveUserSelecion(basewrd : string; selwrd : string);
    procedure btnSettingsClick(Sender: TObject);
    procedure lstFilesKeyPress(Sender: TObject; var Key: Char);
    procedure btnCopyClick(Sender: TObject);
    procedure btnPasteClick(Sender: TObject);
    procedure btnCutClick(Sender: TObject);
    procedure btnUndoClick(Sender: TObject);
    procedure btnRedoClick(Sender: TObject);
    procedure MRUListInsertion(str : string; mru_file : string);
    procedure FormShow(Sender: TObject);
    procedure startup_winTimer(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    file_list : TStringList;
    is_busy : Boolean;
  public
    scan_dir : Boolean;
    app_settings : TAppSettings;
  end;

PFileDataSet = ^TFileDataSet;
TFileDataSet = record
	filename : string;
	dir : string;
end;

const general_err_txt = 'Critical error has been occurred!'#13;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

function TfmMain.Root : string;
begin
  result := IncludeTrailingBackslash(ExtractFilePath(ParamStr(0)));
end;

function TfmMain.LoadAppSettings() : TAppSettings;
var reg1 : TRegIniFile;
begin
  try
    reg1 := TRegIniFIle.Create();
    reg1.RootKey := HKEY_CURRENT_USER;
    result.active_dictionary_id := reg1.ReadInteger('\Software\SpellCode','ActiveDictionary',0);
    result.programming_lan_id := reg1.ReadInteger('\Software\SpellCode','ProgrammingLanguage',0);
    result.show_startup_win := reg1.ReadBool('\Software\SpellCode','ShowStartupScrn',true);
    result.dict_select_message := reg1.ReadBool('\Software\SpellCode','SelectLangMessage',false);
    freeandnil(reg1);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.SaveAppSettings(settings_ : TAppSettings);
var reg1 : TRegIniFile;
begin
  try
    reg1 := TRegINiFile.Create;
    reg1.RootKey := HKEY_CURRENT_USER;
    reg1.WriteInteger('\Software\SpellCode','ActiveDictionary',settings_.active_dictionary_id);
    reg1.WriteInteger('\Software\SpellCode','ProgrammingLanguage',settings_.programming_lan_id);
    reg1.WriteBool('\Software\SpellCode','ShowStartupScrn',settings_.show_startup_win);
    reg1.WriteBool('\Software\SpellCode','SelectLangMessage',settings_.dict_select_message);
    freeandnil(reg1);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

function TfmMain.GetLanIndex(fname : string) : integer;
var ini1 : TIniFile; tsec_buf, ext_list : TStringList; extpos, pos : integer;
begin
  try
    result := -1;
    ini1 := TIniFile.Create(root+'spellcode_lang.ini');
    tsec_buf := TStringList.Create;
    ext_list := TStringList.Create;
    ini1.ReadSections(tsec_buf);
    for pos := 0 to tsec_buf.Count - 1 do begin
      StrToStrings(ini1.ReadString(tsec_buf.Strings[pos],'ext',''),';',ext_list);
      for extpos := 0 to ext_list.Count - 1 do begin
        if (lowercase('.'+trim(ext_list.Strings[extpos])) = Lowercase(ExtractFileExt(fname))) then begin
          result := ini1.ReadInteger(tsec_buf.Strings[pos],'index',-1);
          break;
        end;
      end;
      if result <> -1 then
        break;
    end;
    freeandnil(tsec_buf);
    freeandnil(ext_list);
    freeandnil(ini1);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

function TfmMain.GetSyntaxHighlighter(filename : string) : TSynCustomHighlighter;
begin
  result := nil;
  case GetLanIndex(filename) of
    0 : result := self.hyCPP;
    1 : result := self.hyCS;
    2 : result := self.hyJava;
    3 : result := self.hyPas;
    4 : result := self.hyPerl;
    5 : result := self.hyPHP;
    6 : result := self.hyPython;
    7 : result := self.hyRuby;
    8 : result := self.hyTCL;
    9 : result := self.hyVB;
    10 : result := self.hyXML;
  end;
end;

procedure TfmMain.GetSupportFileExt(file_type : string; ext_list : TStringList);
var ini1 : TIniFile;
begin
  try
		ini1 := TIniFile.Create(root+'spellcode_lang.ini');
		StrToStrings(ini1.ReadString(file_type,'ext',''),';',ext_list);
		freeandnil(ini1);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.BuildScanFileList(mainnode : TTreeNode; slist : TStringList; strdir : string; ext_list : TStringList);
var found_ : boolean; rec_ : TSearchRec; dir_list : TStringList;  pos : integer; _fdata : PFileDataSet;

    function is_valid_file_type(fname : string): Boolean;
    var f_pos : integer;
    begin
      result := false;
      for f_pos := 0 to ext_list.Count - 1 do
        result := result or ('.'+lowercase(trim(ext_list.Strings[f_pos])) = lowercase(ExtractFileExt(fname)));
    end;

begin
  try
    strdir := IncludeTrailingBackslash(strdir);
    found_ := (FindFirst(strdir+'*.*',(faAnyFile-faDirectory),rec_)=0);
    while found_ do begin
      if (is_valid_file_type(rec_.Name)) then begin
        slist.Add(strdir+rec_.Name);
		    New(_fdata);
		    _fdata^.filename := strdir+rec_.Name;
		    _fdata^.dir :=  strdir;
        lstFiles.Items.AddChild(mainnode,rec_.Name).Data := _fdata;
      end;
      found_ := (FindNext(rec_)=0);
    end;
    FindClose(rec_);
    dir_list := TStringList.Create;
    found_ := (FindFirst(strdir+'*.*',faAnyFile, rec_)=0);
    while found_ do begin
      if(((rec_.Attr and faDirectory)<>0) and (rec_.Name[1] <> '.')) then
        dir_list.Add(strdir+rec_.Name);
      found_ := (FindNext(rec_)=0);
    end;
    FindClose(rec_);
    if dir_list.count > 0 then begin
      if (scan_dir or (MessageBox(handle,'This project have sub directories, do you wants to process them for source codes?',PChar(Application.Title), MB_YESNO + MB_ICONQUESTION) = 6)) then begin
        scan_dir := true;
				for pos := 0 to dir_list.Count - 1 do begin
          slist.Add('['+dir_list.Strings[pos]+']');
					BuildScanFileList(lstFiles.Items.AddChild(mainnode, ExtractFileName(dir_list.Strings[pos])),slist, dir_list.Strings[pos], ext_list);
        end;
			end;
    end;
    freeandnil(dir_list);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.btnBrowseClick(Sender: TObject);
var fext_list : TStringList;
begin
  try
    PopBusyToUser;
    if ((Sender = btnRescan) or (dirbrowse.Execute)) then begin
      scan_dir := false;
		  status_panel.Panels[0].Text := 'Scanning the directory contents...';
      fext_list := TStringList.Create;
		  GetSupportFileExt(cmbLanguage.Items.Strings[cmbLanguage.ItemIndex], fext_list);
		  lstFiles.Items.Clear;
      BuildScanFileList(lstFiles.Items.AddChild(nil, ExtractFileName(ExcludeTrailingBackslash(dirbrowse.Directory))), file_list, dirbrowse.Directory, fext_list);
		  freeandnil(fext_list);
      status_panel.Panels[0].Text := '';
      self.btnRescan.Enabled := true;
      lstFiles.Items.GetFirstNode.Expand(false);
      Clear1Click(sender);
      MRUListInsertion(dirbrowse.Directory,root+'spellcode.mru');
      if not app_settings.dict_select_message then begin
        qhelp.ActivateHint(btnLanguage,'Before continue with the SpellCode make sure to select appropriate dictionary from here.','Select the dictionary from here...',6000);
        app_settings.dict_select_message := true;
      end;
    end;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.FormCreate(Sender: TObject);
begin
  try
    is_busy := true;
    LoadAspell;
    file_list := TStringList.Create;
    app_settings := LoadAppSettings;
    cmbLanguage.ItemIndex := app_settings.programming_lan_id;
    LoadASpellDict;
    is_busy := false;
    botbar1.Brush.Color := BrightColor(Panel6.Color,7);
    botbar1.Pen.Color := DarkColor(Panel6.Color,10);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.lstFilesDblClick(Sender: TObject);
begin
  try
    if ((file_list.Count = 0)or(lstFiles.Selected.Data = nil)) then
      exit;
    if not CloseEditorSession then
      exit;
    if (GetAsyncKeyState(VK_SHIFT)<>0) then begin
      ShellExecute(handle,'open',pchar(PFileDataSet(lstFiles.Selected.Data)^.filename),'','',SW_SHOW);
      exit;
    end else if (GetAsyncKeyState(VK_CONTROL)<>0) then begin
      ShellExecute(handle,'open','explorer.exe',pchar('/select, "'+PFileDataSet(lstFiles.Selected.Data)^.filename+'"'),'',SW_SHOW);
      exit;
    end;
    editor1.Lines.LoadFromFile(PFileDataSet(lstFiles.Selected.Data)^.filename);
    assign_project_icon(PFileDataSet(lstFiles.Selected.Data)^.filename);
    editor1.Highlighter := GetSyntaxHighlighter(extractfilename(PFileDataSet(lstFiles.Selected.Data)^.filename));
    lblFilename.Caption := extractfilename(PFileDataSet(lstFiles.Selected.Data)^.filename);
    lblFilename.Font.Style := lblFilename.Font.Style + [fsBold];
    lblFilename.Hint := PFileDataSet(lstFiles.Selected.Data)^.filename;
    editor1.Enabled := true;
    editor1.Modified := false;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.LoadASpellDict();
var config_ : AspellConfig; dict_list : AspellDictInfoList; dinfo_set : AspellDictInfoEnumeration;
dict_num : integer; act_dict_info : TAspellDictInfo; mnu_ : TMenuItem;
begin
  try
    config_ := new_aspell_config;
    dict_list := get_aspell_dict_info_list(config_);
    delete_aspell_config(config_);
    dinfo_set := aspell_dict_info_list_elements(dict_list);
    dict_num := 0;
    repeat
      inc(dict_num);
      act_dict_info := (aspell_dict_info_enumeration_next(dinfo_set)^);
      mnu_ := TMenuItem.Create(self);
      mnu_.Hint := act_dict_info.name;
      mnu_.Caption := GetDictDisplayName(mnu_.Hint);
      mnu_.Tag := dict_num;
      mnu_.Checked := (dict_num = app_settings.active_dictionary_id);
      if mnu_.Checked then
        LanguageMenuClick(mnu_);
      mnu_.OnClick := LanguageMenuClick;
      popDict.Items.Add(mnu_);
    until (aspell_dict_info_enumeration_at_end(dinfo_set) <> 0);
    delete_aspell_dict_info_enumeration(dinfo_set);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.LanguageMenuClick(Sender : TObject);
var mpos : Integer;
begin
  try
    btnLanguage.Caption := ' '+TMenuItem(Sender).Caption;
    btnLanguage.Tag := TMenuItem(Sender).Tag;
    for mpos := 0 to popDict.Items.Count - 1 do
      popDict.Items.Items[mpos].Checked := (TMenuItem(Sender).MenuIndex = mpos);
    btnLanguage.Hint := 'Active dictionary is "' + StringReplace(TMenuItem(Sender).Caption,'&','',[rfReplaceAll])+'"';
    app_settings.active_dictionary_id := TMenuItem(Sender).Tag;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

function TfmMain.GetDictDisplayName(did : string) : string;
var ini1 : TINiFile;
begin
  try
    ini1 := TINiFile.Create(root+'spellcode_dict.ini');
    result := ini1.ReadString(did,'name',did);
    freeandnil(ini1);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  try
    SaveAppSettings(app_settings);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.cmbLanguageChange(Sender: TObject);
begin
  app_settings.programming_lan_id := cmbLanguage.ItemIndex;
end;

function TfmMain.GetActiveDict : string;
var pos : integer;
begin
  for pos := 0 to self.popDict.Items.Count - 1 do begin
    if (pos = btnLanguage.Tag) then begin
      result := popDict.Items.Items[pos].hint;
      break;
    end;
  end;
end;

procedure TfmMain.SentzAnalyzer(str : string; wrd_list : TStringList; pos_list : TStringList);
var pos : integer;

  procedure StrToStringsEx2(S : string; const List: TStrings; const IndexList : TStrings);
  var spos, pos, sep_count : integer;
  begin
    list.Clear;
    IndexList.Clear;
    if s = '' then
      EXIT;
    spos := 0;
    sep_count := 0;
    for pos := 0 to Length(s) do begin
      if s[pos] = ' ' then begin
        List.Add(StrMid(S,spos,(pos-spos)));
        if (StrMid(S,spos,(pos-spos)) = ' ') then
          inc(sep_count)
        else
          sep_count := 0;
        IndexList.Add(inttostr(spos+sep_count+1));
        spos := pos;
      end;
    end;
    IndexList.Add(inttostr(spos+sep_count+1));
    List.Add(StrMid(S,spos,(pos-spos)));
  end;

begin
  try
    wrd_list.Clear;
    if (trim(str)='') then
      exit;
    StrReplace(str,'//','  ',[rfReplaceAll]);
    StrReplace(str,'/*','  ',[rfReplaceAll]);
    StrReplace(str,'*/','  ',[rfReplaceAll]);
    StrReplace(str,'{',' ',[rfReplaceAll]);
    StrReplace(str,'}',' ',[rfReplaceAll]);
    StrReplace(str,'<!--','    ',[rfReplaceAll]);
    StrReplace(str,'-->','   ',[rfReplaceAll]);
    StrReplace(str,'''',' ',[rfReplaceAll]);
    StrReplace(str,'"',' ',[rfReplaceAll]);
    StrReplace(str,'\',' ',[rfReplaceAll]);
    StrReplace(str,'/',' ',[rfReplaceAll]);
    StrReplace(str,'.',' ',[rfReplaceAll]);
    StrReplace(str,',',' ',[rfReplaceAll]);
    StrReplace(str,':',' ',[rfReplaceAll]);
    StrReplace(str,';',' ',[rfReplaceAll]);
    StrToStringsEx2(str,wrd_list, pos_list);
    pos := 0;
    while(pos < wrd_list.Count) do begin
      if ((trim(wrd_list.Strings[pos]) = '') or (Length(trim(wrd_list.Strings[pos])) < 2)) then begin
        wrd_list.Delete(pos);
        pos_list.Delete(pos);
        Continue;
      end;
      if(trim(wrd_list.Strings[pos]) = UpperCase(trim(wrd_list.Strings[pos]))) then begin
        wrd_list.Delete(pos);
        pos_list.Delete(pos);
        Continue;
      end;
      if (not StrIsAlpha(trim(wrd_list.Strings[pos]))) then begin
        wrd_list.Delete(pos);
        pos_list.Delete(pos);
        Continue;
      end else
        inc(pos);
    end;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

function TfmMain.is_valid_token(tkid : integer; lan_id : integer) : Boolean;
begin
  case lan_id of
    0 : result := ((tkid = Ord(SynHighlighterCpp.tkComment)) or (tkid = ord(SynHighlightercpp.tkString)));
    1 : result := ((tkid = Ord(SynHighlighterCS.tkComment)) or (tkid = Ord(SynHighlighterCS.tkString)));
    2 : result := ((tkid = Ord(SynHighlighterJava.tkComment)) or (tkid = Ord(SynHighlighterJava.tkString)));
    3 : result := ((tkid = Ord(SynHighlighterPas.tkComment)) or (tkid = Ord(SynHighlighterPas.tkString)));
    4 : result := ((tkid = Ord(SynHighlighterPerl.tkComment)) or (tkid = Ord(SynHighlighterPerl.tkString)));
    5 : result := ((tkid = Ord(SynHighlighterPhp.tkComment)) or (tkid = Ord(SynHighlighterPhp.tkString)));
    6 : result := ((tkid = Ord(SynHighlighterPython.tkComment)) or (tkid = Ord(SynHighlighterPython.tkString)));
    7 : result := ((tkid = Ord(SynHighlighterruby.tkComment)) or (tkid = Ord(SynHighlighterruby.tkString)));
    8 : result := ((tkid = Ord(SynHighlighterTCLtk.tkComment)) or (tkid = Ord(SynHighlighterTCLtk.tkString)));
    9 : result := ((tkid = Ord(SynHighlightervb.tkComment)) or (tkid = Ord(SynHighlightervb.tkString)));
    10: result := (tkid = Ord(SynHighlighterxml.tkComment));
  end;
end;

procedure TfmMain.btnCheckSpellClick(Sender: TObject);
var dconf : AspellConfig; is_dict_error : AspellCanHaveError;
spchk : AspellSpeller; posline, wscan, ret : integer; word_list, pos_list : TStringList;
sug_list : AspellWordList; sug_elements : AspellStringEnumeration; delta : integer;
sug_word : PChar; hint_win : TfmHint; ret_code : integer;
begin
  try
    if (trim(editor1.Text) = '') then begin
      if msgview1.Items.Count > 0 then
        qhelp.ActivateHint(msgview1,'Please select some source file from here or '#13'from the project tree view and then start the spell checking.','Need to have some source file...',6000)
      else if lstFiles.Items.Count > 0 then
        qhelp.ActivateHint(lstFiles,'Please double click and select some source file from this project tree view and then start spell checking.','Need to have some source file...',6000)
      else
        qhelp.ActivateHint(btnBrowse,'Before start with the spell checking you need to open some project from here.','Need to have some source file...',6000);
      exit;
    end;
    is_busy := true;
    ret_code := mrNone;
    hint_win := TfmHint.Create(self);
    dconf := new_aspell_config();
    aspell_config_replace(dconf,'lang',PChar(GetActiveDict));
    is_dict_error := new_aspell_speller(dconf);
    spchk := to_aspell_speller(is_dict_error);
    delete_aspell_config(dconf);
    posline := 0;
    word_list := TStringList.Create;
    pos_list :=  TStringList.Create;
    while posline <= editor1.Lines.Count do begin
      editor1.Highlighter.SetLine(editor1.Lines.Strings[posline],posline);
      delta := 0;
      while not editor1.Highlighter.GetEol do begin
        if is_valid_token(editor1.Highlighter.GetTokenKind, GetLanIndex(lblFileName.Caption)) then begin
          SentzAnalyzer(editor1.Highlighter.GetToken,word_list, pos_list);
          if word_list.Count > 0 then begin
            wscan := 0;
            while(wscan < word_list.Count) do begin
              ret := aspell_speller_check(spchk, PChar(trim(word_list.Strings[wscan])), Length(trim(word_list.Strings[wscan])));
              if (ret <> 1) then begin
                sug_list := aspell_speller_suggest(spchk, PChar(trim(word_list.Strings[wscan])), Length(trim(word_list.Strings[wscan])));
                sug_elements := aspell_word_list_elements(sug_list);
                hint_win.lblWord.Caption := trim(word_list.Strings[wscan]);
                hint_win.lstSugs.Items.Clear;
                repeat
                  sug_word := aspell_string_enumeration_next(sug_elements);
                  if (sug_word <> nil) then
                    hint_win.lstSugs.Items.Add(sug_word);
                until (sug_word = nil);
                delete_aspell_string_enumeration(sug_elements);
                if hint_win.lstSugs.Items.Count > 0 then
                  hint_win.lstSugs.ItemIndex := 0;
                editor1.GotoLineAndCenter(posline+1);
                editor1.CaretX := StrToInt(pos_list.Strings[wscan])+ delta;
                editor1.SelLength := Length(trim(word_list.Strings[wscan]));
                hint_win.Top := (((editor1.CaretY - editor1.TopLine)+1) * editor1.LineHeight)+top+editor1.Top+editor1.LineHeight+GetSystemMetrics(SM_CYCAPTION);
                hint_win.Left := left + Panel3.Left + Panel4.Left + editor1.Left+editor1.RowColumnToPixels(editor1.DisplayXY).X;
                ret_code := hint_win.ShowModal;
                if (ret_code = mrNo) then
                  break
                else if(ret_code = mrYes) then
                  aspell_speller_add_to_personal(spchk,PChar(trim(word_list.Strings[wscan])), Length(trim(word_list.Strings[wscan])))
                else if((ret_code = mrOK) and (hint_win.lstSugs.itemindex > -1)) then begin
                  delta := delta + (Length(hint_win.lstSugs.Items.Strings[hint_win.lstSugs.itemindex]) - Length(trim(word_list.Strings[wscan])));
                  editor1.SelText := hint_win.lstSugs.Items.Strings[hint_win.lstSugs.itemindex];
                  aspell_speller_store_replacement(spchk, PChar(trim(word_list.Strings[wscan])), Length(trim(word_list.Strings[wscan])), pchar(hint_win.lstSugs.Items.Strings[hint_win.lstSugs.itemindex]), Length(hint_win.lstSugs.Items.Strings[hint_win.lstSugs.itemindex]));
                end else if(ret_code = mrYesToAll) then
                  aspell_speller_add_to_session(spchk,PChar(trim(word_list.Strings[wscan])), Length(trim(word_list.Strings[wscan])));
              end;
              inc(wscan);
            end;
            if (ret_code = mrNo) then
              break;
          end;
        end;
        if (ret_code = mrNo) then
          break;
        delta := delta+Length(editor1.Highlighter.GetToken);
        editor1.Highlighter.Next;
      end;
      inc(posline);
      if (ret_code = mrNo) then
        break;
    end;
    aspell_speller_save_all_word_lists(spchk);
    delete_aspell_speller(spchk);
    freeandnil(word_list);
    freeandnil(hint_win);
    freeandnil(pos_list);
    editor1.SelLength := 0;
    is_busy := false;
    MessageBox(handle,'The spelling check is complete.',PChar(Application.Title),MB_OK+MB_ICONINFORMATION);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.btnScanClick(Sender: TObject);
var tot_count, err_cnt : integer; veditor : TSynEdit; dconf : AspellConfig;
is_dict_error : AspellCanHaveError; spchk : AspellSpeller; flan_index : integer;

  function GetTotalNodeCount(pnode : TTreeNode) : integer;
  var pos : integer;
  begin
    result := 0;
    for pos := 0 to pnode.Count - 1 do begin
      if pnode.Item[pos].HasChildren then
        result := result + GetTotalNodeCount(pnode.Item[pos]);
    end;
    result := result + pos;
  end;

  procedure MakeLog(mnode : TTreeNode; filename : string; row_ : integer; col_ : integer; err_word : string; token_kind : integer);
  var tk_type : string;

    function is_comment(state : Boolean) : string;
    begin
      if state then
        result := 'Comment'
      else
        result := 'String';
    end;

  begin
    tk_type := 'Unknown';
    case flan_index of
      0 : tk_type := is_comment(Ord(SynHighlighterCpp.tkComment) = token_kind);
      1 : tk_type := is_comment(Ord(SynHighlighterCS.tkComment) = token_kind);
      2 : tk_type := is_comment(Ord(SynHighlighterJava.tkComment) = token_kind);
      3 : tk_type := is_comment(Ord(SynHighlighterPas.tkComment) = token_kind);
      4 : tk_type := is_comment(Ord(SynHighlighterPerl.tkComment) = token_kind);
      5 : tk_type := is_comment(Ord(SynHighlighterPhp.tkComment) = token_kind);
      6 : tk_type := is_comment(Ord(SynHighlighterPython.tkComment) = token_kind);
      7 : tk_type := is_comment(Ord(SynHighlighterRuby.tkComment) = token_kind);
      8 : tk_type := is_comment(Ord(SynHighlighterTclTk.tkComment) = token_kind);
      9 : tk_type := is_comment(Ord(SynHighlighterVb.tkComment) = token_kind);
      10: tk_type := 'Comment';
    end;
    msgview1.Items.AddChild(mnode,extractfilename(filename)+';'+err_word + ';'+format('%d: %d',[row_, col_])+';'+tk_type);
  end;

  function ScanFile(filename : string) : integer;
  var posline, delta, wscan, ret : integer; word_list, pos_list : TStringList; pnode : TTreeNode;
  begin
    result := 0;
    status_panel.Panels.Items[0].Text := 'Scanning '+StringReplace(filename, ExcludeTrailingBackslash(dirbrowse.Directory),'',[rfIgnoreCase]) + '...';
    Application.ProcessMessages;
    pnode := msgview1.Items.Add(nil,StringReplace(filename, ExcludeTrailingBackslash(dirbrowse.Directory),'',[rfIgnoreCase]));
    self.Repaint;
    veditor.Lines.LoadFromFile(filename);
    veditor.Highlighter := GetSyntaxHighlighter(extractfilename(filename));
    prgLevel.Position := prgLevel.Position + 1;
    if (trim(veditor.Lines.Text) = '') then
      exit;
    posline := 0;
    word_list := TStringList.Create;
    pos_list :=  TStringList.Create;
    while posline <= veditor.Lines.Count do begin
      veditor.Highlighter.SetLine(veditor.Lines.Strings[posline],posline);
      delta := 0;
      while not veditor.Highlighter.GetEol do begin
        flan_index := GetLanIndex(extractfilename(filename));
        if is_valid_token(veditor.Highlighter.GetTokenKind, flan_index) then begin
          SentzAnalyzer(veditor.Highlighter.GetToken,word_list, pos_list);
          if word_list.Count > 0 then begin
            wscan := 0;
            while(wscan < word_list.Count) do begin
              ret := aspell_speller_check(spchk, PChar(trim(word_list.Strings[wscan])), Length(trim(word_list.Strings[wscan])));
              if (ret <> 1) then begin
                MakeLog(pnode, filename,posline+1,delta+StrToInt(pos_list.Strings[wscan]),trim(word_list.Strings[wscan]),veditor.Highlighter.GetTokenKind);
                result := result + 1;
              end;
              inc(wscan);
            end;
          end;
        end;
        delta := delta+Length(veditor.Highlighter.GetToken);
        veditor.Highlighter.Next;
      end;
      inc(posline);
    end;
    freeandnil(word_list);
    freeandnil(pos_list);
  end;

  function ScanTreeNode(pnode : TTreeNode) : integer;
  var _pos : integer;
  begin
    result := 0;
    for _pos := 0 to pnode.Count - 1 do begin
      if pnode.Item[_pos].HasChildren then
        result := result + ScanTreeNode(pnode.Item[_pos])
      else begin
        if ((pnode.Item[_pos].Data <> nil) and (fileexists(PFileDataSet(pnode.Item[_pos].Data)^.filename))) then
          result := result + ScanFile(PFileDataSet(pnode.Item[_pos].Data)^.filename);
      end;
    end;
  end;

begin
  try
    PopBusyToUser;
    if lstFiles.Items.Count = 0 then begin
      qhelp.ActivateHint(btnBrowse,'Please select some directory from here and then press '#13'the "Check" button to get the list of spelling mistakes. ','Select some location to scan...',6000);
      exit;
    end;
    InterfaceState(false);
    prgLevel.Maximum := GetTotalNodeCount(lstFiles.Items.Item[0]);
    msgview1.Items.Clear;
    veditor := TSynEdit.Create(nil);
    dconf := new_aspell_config();
    aspell_config_replace(dconf,'lang',PChar(GetActiveDict));
    is_dict_error := new_aspell_speller(dconf);
    spchk := to_aspell_speller(is_dict_error);
    delete_aspell_config(dconf);
    err_cnt := ScanTreeNode(lstFiles.Items.Item[0]);
    delete_aspell_speller(spchk);
    freeandnil(veditor);
    prgLevel.Position := 0;
    status_panel.Panels.Items[0].Text := '';
    if msgview1.Items.Count > 0 then
      msgview1.Items.Item[0].Selected := true;
    InterfaceState(true);
    setup_taskbar_button_flash;
    MessageBox(handle,pchar(format('Scan %d project source file(s) and found %d spelling mistake(s).',[prgLevel.Maximum, err_cnt])),PChar(Application.Title),MB_OK+MB_ICONINFORMATION);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.msgview1DblClick(Sender: TObject);
var strlist : TStringList; wrd : string;
begin
  try
    if ((msgview1.Items.Count = 0) or (msgview1.Selected.HasChildren)) then
      exit;
    if (StrFind(';',msgview1.Selected.Text) = 0) then
      exit;
    if not CloseEditorSession then
      exit;
    editor1.Lines.LoadFromFile(ExcludeTrailingBackslash(dirbrowse.Directory) + msgview1.Selected.Parent.Text);
    assign_project_icon(ExcludeTrailingBackslash(dirbrowse.Directory) + msgview1.Selected.Parent.Text);
    lblFilename.Hint := ExcludeTrailingBackslash(dirbrowse.Directory) + msgview1.Selected.Parent.Text;
    editor1.Highlighter := GetSyntaxHighlighter(extractfilename(lblFilename.Hint));
    editor1.Enabled := true;
    lblFilename.Font.Style := lblFilename.Font.Style + [fsBold];
    lblFilename.Caption := extractfilename(lblFilename.Hint);
    strlist := TStringList.Create;
    StrToStrings(msgview1.Selected.Text,';',strlist,true);
    wrd := strlist.Strings[1];
    StrToStrings(strlist.Strings[2],':',strlist,true);
    editor1.GotoLineAndCenter(StrToIntDef(trim(strlist.Strings[0]),0));
    editor1.CaretX := StrToIntDef(trim(strlist.Strings[1]),0);
    editor1.SelLength := Length(wrd);
    editor1.SetFocus;
    editor1.Modified := false;
    freeandnil(strlist);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.InterfaceState(state : Boolean);
begin
  fmMain.Invalidate;
  is_busy := not state;
  lstFiles.Enabled := state;
  cmbLanguage.Enabled := lstFiles.Enabled;
  editor1.Enabled := lstFiles.Enabled;
  msgview1.Enabled := lstFiles.Enabled;
  btnLanguage.Enabled := lstFiles.Enabled;
  fmMain.Update;
end;

function TfmMain.PopBusyToUser() : boolean;
begin
  result := is_busy;
  if is_busy then
    MessageBox(handle,'SpellCode is busy with some data processing operation, please retry in a few minutes.',PChar(Application.Title),MB_OK+MB_ICONWARNING);
end;

procedure TfmMain.btnSaveClick(Sender: TObject);
var dummy_num : integer;
begin
  try
    dummy_num := 0;
    if (editor1.Enabled and editor1.Modified and (lblFilename.Hint <> '') and is_file_attrib_clr(lblFilename.Hint,false,dummy_num)) then begin
      editor1.Lines.SaveToFile(lblFilename.Hint);
      editor1.Modified := false;
    end;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

function TfmMain.CloseEditorSession() : Boolean;
var ret_cd : integer;
begin
  try
    if (editor1.Enabled and (editor1.Modified)) then begin
      result := true;
      ret_cd := MessageBox(handle,pchar('Do you want to save the changes to '+extractfilename(lblFilename.Hint)+'?'),PChar(Application.Title),MB_YESNOCANCEL+MB_ICONQUESTION);
      if ret_cd = mrCancel then
        result := false
      else if ret_cd = mrYes then
        btnSaveClick(nil);
    end else
      result := true;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := CloseEditorSession;
end;

procedure TfmMain.editor1StatusChange(Sender: TObject;
  Changes: TSynStatusChanges);
begin
  if editor1.Modified then
    status_panel.Panels.Items[3].Text := 'Modified'
  else
    status_panel.Panels.Items[3].Text := '';
end;

procedure TfmMain.btnAboutClick(Sender: TObject);
var abf_frm : TfmAbout; base_ver : TJclFileVersionInfo;
begin
  try
    abf_frm := TfmAbout.Create(self);
    base_ver := TJclFileVersionInfo.Create(ParamStr(0));
    abf_frm.lblversion.Caption := 'Version '+base_ver.FileVersion;
    freeandnil(base_ver);
    abf_frm.ShowModal;
    freeandnil(abf_frm);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.msgview1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  try
    Locatethemistake1.Enabled := false;
    Getthenumberofoccurrences1.Enabled := Locatethemistake1.Enabled;
    if ((msgview1.GetNodeAt(X,Y) <> nil) and (not msgview1.GetNodeAt(X,Y).HasChildren) and (StrFind(';',msgview1.GetNodeAt(X,Y).Text)>0)) then begin
      msgview1.GetNodeAt(X,Y).Selected := true;
      Locatethemistake1.Enabled := true;
      Getthenumberofoccurrences1.Enabled := Locatethemistake1.Enabled;
    end else if (msgview1.GetNodeAt(X,Y) <> nil) then begin
      Getthenumberofoccurrences1.Enabled := true;
      msgview1.GetNodeAt(X,Y).Selected := true;
    end; 
    Ignorealloccurrences1.Enabled := Locatethemistake1.Enabled;
    Replacewith1.Enabled := Locatethemistake1.Enabled;
    Addtodictionary1.Visible := Locatethemistake1.Enabled;
    Copy1.Enabled := Locatethemistake1.Enabled;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.inentireproject1Click(Sender: TObject);
var _path, str, tstr : string; strlist : TStringList; pos, rem_count : integer;
begin
  try
    rem_count := 0;
    strlist := TStringList.Create;
    StrToStrings(msgview1.Selected.Text,';',strlist);
    tstr := trim(strlist.Strings[1]);
    str := lowercase(tstr);
    pos := msgview1.Items.Count - 1;
    if TMenuItem(Sender).Tag = 1 then
      _path := trim(lowercase(msgview1.Selected.Parent.Text));
    InterfaceState(false);
    while pos >= 0 do begin
      if (StrFind(';',msgview1.Items.Item[pos].Text) > 0) then begin
        StrToStrings(msgview1.Items.Item[pos].Text,';',strlist);
        if (lowercase(trim(strlist.Strings[1])) = str) then begin
          if ((TMenuItem(Sender).Tag = 2) or (TMenuItem(Sender).Tag = 0) or ( (TMenuItem(Sender).Tag = 1) and (_path = trim(lowercase(msgview1.Items.Item[pos].Parent.Text))))) then begin
            msgview1.Items.Item[pos].Delete;
            inc(rem_count);
          end;
        end;
      end;
      dec(pos);
    end;
    InterfaceState(true);
    if TMenuItem(Sender).Tag <> 2 then
      MessageBox(handle,pchar('Ignore the word "'+trim(tstr)+'", which occurred in '+inttostr(rem_count)+' location(s).'),pchar(Application.Title),MB_OK+MB_ICONINFORMATION);
    freeandnil(strlist);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.Addtodictionary1Click(Sender: TObject);
var str : string; strlist : TStringList; dconf : AspellConfig; is_dict_error : AspellCanHaveError;
spchk : AspellSpeller;
begin
  try
    strlist := TStringList.Create;
    StrToStrings(msgview1.Selected.Text,';',strlist);
    str := trim(strlist.Strings[1]);
    freeandnil(strlist);
    dconf := new_aspell_config();
    aspell_config_replace(dconf,'lang',PChar(GetActiveDict));
    is_dict_error := new_aspell_speller(dconf);
    spchk := to_aspell_speller(is_dict_error);
    delete_aspell_config(dconf);
    aspell_speller_add_to_personal(spchk,PChar(str), Length(str));
    aspell_speller_save_all_word_lists(spchk);
    delete_aspell_speller(spchk);
    inentireproject1Click(sender);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.Clear1Click(Sender: TObject);
var do_clr : Boolean;
begin
  try
    if (msgview1.Items.Count = 0) then
      exit;
    do_clr := true;
    if TMenuItem(Sender).Tag = 1 then
      do_clr := (MessageBox(Handle,'Are you sure do you want to clear the entire spell checking result log?',PChar(Application.Title),MB_YESNO+MB_ICONQUESTION) = 6);
    if do_clr then
      msgview1.Items.Clear;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.msgview_popupPopup(Sender: TObject);
var str : string; strlist : TStringList; dconf : AspellConfig; sug_list : AspellWordList; 
is_dict_error : AspellCanHaveError; spchk : AspellSpeller; cnt_ : integer;
sug_elements : AspellStringEnumeration; sug_word : PChar; mnu_item : TMenuItem;

  procedure ClearPopupMenuSugList();
  begin
    while (Replacewith1.Count > 0) do
      Replacewith1.Items[0].Free;
  end;

begin
  try
    if not Locatethemistake1.Enabled then
      exit;
    strlist := TStringList.Create;
    StrToStrings(msgview1.Selected.Text,';',strlist);
    str := trim(strlist.Strings[1]);
    freeandnil(strlist);
    Addtodictionary1.Caption := 'Add "'+str+'" to dictionary ';
    ClearPopupMenuSugList;
    dconf := new_aspell_config();
    aspell_config_replace(dconf,'lang',PChar(GetActiveDict));
    is_dict_error := new_aspell_speller(dconf);
    spchk := to_aspell_speller(is_dict_error);
    delete_aspell_config(dconf);
    sug_list := aspell_speller_suggest(spchk, PChar(str), Length(str));
    sug_elements := aspell_word_list_elements(sug_list);
    cnt_ := 0;
    repeat
      sug_word := aspell_string_enumeration_next(sug_elements);
      if (sug_word <> nil) then begin
        mnu_item := TMenuItem.Create(nil);
        mnu_item.Caption := sug_word;
        mnu_item.AutoHotkeys := maAutomatic;
        mnu_item.Hint := str;
        mnu_item.OnClick := replace_word_menu_click;
        Replacewith1.Add(mnu_item);
        inc(cnt_);
      end;
    until ((sug_word = nil) or (cnt_ = 10));
    Replacewith1.Enabled := (cnt_ > 0);
    delete_aspell_string_enumeration(sug_elements);
    aspell_speller_save_all_word_lists(spchk);
    delete_aspell_speller(spchk);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

function TfmMain.is_file_attrib_clr(filename : string; show_all_btns : boolean; var old_status : integer) : Boolean;
var attrib_, rslt: integer; is_chg : boolean; btncode : TMsgDlgButtons;
begin
  try
    result := false;
    attrib_ := FileGetAttr(filename);
    if ((attrib_ and faReadOnly) <> 0) then begin
      is_chg := false;
      if ((old_status = 0) or (old_status = mrYes) or (old_status = mrNo)) then begin
        if show_all_btns then
          btncode := [mbYes,mbNo,mbYesToAll,mbNoToAll]
        else
          btncode := [mbYes,mbNo];
        rslt  := MessageDlg('The file "'+extractfilename(filename)+'" is marked as read-only file. Are you sure do you want to edit this file?',mtConfirmation,btncode,0);
        is_chg := ((rslt = mrYes) or (rslt = mrYesToAll));
        old_status := rslt;
      end else
        is_chg := (old_status = mrYesToAll);
      if is_chg then begin
        FileSetAttr(filename,(attrib_ - faReadOnly));
        result := true;
      end;
    end else
      result := true;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.replace_word_menu_click(Sender : TObject);
var pos : integer; str : string; strlist, pos_str : TStringList; tmp_edit : TSynEdit;
pline, delta : integer; fname, _tmp : string; rdo_code : integer;
begin
  try
    if not CloseEditorSession then
      exit;
    strlist := TStringList.Create;
    StrToStrings(msgview1.Selected.Text,';',strlist);
    if (MessageBox(Handle,PChar('Replace all the occurrences of "'+trim(strlist.Strings[1])+'" with "'+StringReplace(TMenuItem(Sender).Caption,'&','',[])+'" and recheck the project?'),pchar(Application.Title),MB_YESNO+MB_ICONQUESTION) = 6) then begin
      pos_str := TStringList.Create;
      str := lowercase(strlist.Strings[1]);
      tmp_edit := TSynEdit.Create(nil);
      fname := '';
      pline := 0;
      delta := 0;
      rdo_code := 0;
      InterfaceState(false);
      for pos := 0 to msgview1.Items.Count - 1 do begin
        if StrFind(';',msgview1.Items.Item[pos].Text) > 0 then begin
          StrToStrings(msgview1.Items.Item[pos].Text,';',strlist);
          if(trim(str) = lowercase(trim(strlist.Strings[1]))) then begin
            if (fname <> ExcludeTrailingBackslash(dirbrowse.Directory) + msgview1.Items.Item[pos].Parent.Text) then begin
              fname := ExcludeTrailingBackslash(dirbrowse.Directory) + msgview1.Items.Item[pos].Parent.Text;
              tmp_edit.Lines.LoadFromFile(fname);
              delta := 0;
            end;
            StrToStrings(msgview1.Items.Item[pos].Text,';',pos_str);
            StrToStrings(pos_str.Strings[2], ':', pos_str);
            tmp_edit.CaretY := StrToIntDef(trim(pos_str.Strings[0]),0);
            if pline <> tmp_edit.CaretY then
              delta := 0;
            tmp_edit.CaretX := StrToIntDef(trim(pos_str.Strings[1]),0)+delta;
            delta := delta + (Length(StringReplace(TMenuItem(Sender).Caption,'&','',[])) - Length(trim(str)));
            tmp_edit.SelLength := Length(str);
            pline := StrToIntDef(trim(pos_str.Strings[0]),0);
            if (trim(lowercase(tmp_edit.SelText)) = trim(str)) then
              tmp_edit.SelText := StringReplace(TMenuItem(Sender).Caption,'&','',[]);
            if is_file_attrib_clr(ExcludeTrailingBackslash(dirbrowse.Directory) + msgview1.Items.Item[pos].Parent.Text,true,rdo_code) then
              tmp_edit.Lines.SaveToFile(ExcludeTrailingBackslash(dirbrowse.Directory) + msgview1.Items.Item[pos].Parent.Text);
          end;
        end;
      end;
      SaveUserSelecion(TMenuItem(Sender).Hint, StringReplace(TMenuItem(Sender).Caption,'&','',[]));
      InterfaceState(true);
      btnScanClick(Sender);
      freeandnil(tmp_edit);
      freeandnil(pos_str);
      if (lblFilename.Hint <> '') then begin
        editor1.Lines.LoadFromFile(lblFilename.Hint);
        editor1.Modified := false;
      end;
    end;
    freeandnil(strlist);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.setup_taskbar_button_flash();
var flsh_struct : FLASHWINFO;
begin
  try
    if (lowercase(GetProcessNameFromWnd(GetForegroundWindow)) <> lowercase(Paramstr(0))) then begin
      flsh_struct.cbSize := 20;
      flsh_struct.hwnd := Application.Handle;
      flsh_struct.dwFlags := FLASHW_TIMERNOFG + FLASHW_TRAY;
      flsh_struct.uCount := 10;
      flsh_struct.dwTimeout := 0;
      FlashWindowEx(flsh_struct);
    end;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.SaveAs1Click(Sender: TObject);
var fcont : TStringList; pos : integer; adv_ : string;
begin
  try
    if ((msgview1.Items.Count > 0) and (logsave1.Execute)) then begin
      fcont := TStringList.Create;
      for pos := 0 to msgview1.Items.Count - 1 do begin
        if StrFind(';',msgview1.Items.Item[pos].Text) > 0 then
          adv_ := #9
        else
          adv_ := '';
        fcont.Add(adv_+StringReplace(msgview1.Items.Item[pos].Text,';',#9#9,[rfReplaceAll, rfIgnoreCase]));
      end;
      fcont.SaveToFile(logsave1.FileName);
      freeandnil(fcont);
    end else if (msgview1.Items.Count = 0) then
      MessageBeep(MB_ICONWARNING);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.app_evntsSettingChange(Sender: TObject; Flag: Integer;
  const Section: String; var Result: Integer);
begin
  botbar1.Brush.Color := BrightColor(Panel6.Color,7);
  botbar1.Pen.Color := DarkColor(Panel6.Color,10);
end;

procedure TfmMain.Getthenumberofoccurrences1Click(Sender: TObject);
var fname : string;
begin
  try
    if msgview1.Items.Count = 0 then
      exit;
    if (StrFind(';', msgview1.Selected.Text) = 0) then
      fname := msgview1.Selected.Text
    else
      fname := msgview1.Selected.Parent.Text;
    ShellExecute(0,'open','explorer.exe',pchar('/select, "'+ExcludeTrailingBackslash(dirbrowse.Directory) + fname+'"'),'',SW_SHOW);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.assign_project_icon(filename : string);
var finfo: TSHFileInfo; img1: TImageList;
begin
  try
    img1 := TImageList.Create(nil);
    with img1 do
    begin
      Handle := SHGetFileInfo(pchar(filename), 0, finfo, SizeOf(TSHFileInfo), SHGFI_SMALLICON or SHGFI_SYSICONINDEX );
      ShareImages := True;
    end;
    img1.GetIcon(finfo.iIcon,fileicon.Picture.Icon);
    freeandnil(img1);
    fileicon.Hint := filename;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.fileiconMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  try
    if ((lblFilename.Hint <> '') and (Button = mbRight)) then
      DisplayContextMenu(Panel4.Handle,lblFilename.Hint,Point(X+fileicon.Left+6,Y+fileicon.Top+6));
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.fileiconDblClick(Sender: TObject);
begin
  try
    if (lblFilename.Hint <> '') then
      ShellExecute(0,'open',pchar(lblFilename.Hint),'','',SW_SHOW);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.ReplaceWithSpotWord(Sender: TObject);
begin
  try
    if (Trim(editor1.SelText) = '') then
      exit;
    editor1.seltext := StringReplace(TMenuItem(Sender).Caption,'&','',[rfReplaceAll]);
    SaveUserSelecion(TMenuItem(Sender).Hint,StringReplace(TMenuItem(Sender).Caption,'&','',[rfReplaceAll]));
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.sugmenuPopup(Sender: TObject);
var menu_item : TMenuItem; dconf : AspellConfig; sug_list : AspellWordList;
is_dict_error : AspellCanHaveError; spchk : AspellSpeller; sug_word : PChar;
sug_elements : AspellStringEnumeration; cnt_ : integer; 
begin
  try
    while (sugmenu.Items.Count > 0) do
      sugmenu.items.Items[0].Free;
    if (Trim(editor1.SelText) = '') then
      exit;
    dconf := new_aspell_config();
    aspell_config_replace(dconf,'lang',PChar(GetActiveDict));
    is_dict_error := new_aspell_speller(dconf);
    spchk := to_aspell_speller(is_dict_error);
    delete_aspell_config(dconf);
    sug_list := aspell_speller_suggest(spchk, PChar(editor1.SelText), Length(editor1.SelText));
    sug_elements := aspell_word_list_elements(sug_list);
    cnt_ := 0;
    repeat
      sug_word := aspell_string_enumeration_next(sug_elements);
      if (sug_word <> nil) then begin
        menu_item := TMenuItem.Create(nil);
        menu_item.Caption := sug_word;
        menu_item.Hint := editor1.SelText;
        menu_item.Default := (lowercase(sug_word ) = lowercase(trim(editor1.SelText)));
        if not menu_item.Default then
          menu_item.OnClick := ReplaceWithSpotWord;
        sugmenu.Items.Add(menu_item);
        inc(cnt_);
      end;
    until ((sug_word = nil) or (cnt_ = 15));
    delete_aspell_string_enumeration(sug_elements);
    aspell_speller_save_all_word_lists(spchk);
    delete_aspell_speller(spchk);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.Copy1Click(Sender: TObject);
var strlist : TStringList; cbrd : TClipboard;
begin
  try
    if ((msgview1.Items.Count = 0) or (msgview1.Selected.HasChildren)) then
      exit;
    if (StrFind(';',msgview1.Selected.Text) = 0) then
      exit;
    strlist := TStringList.Create;
    StrToStrings(msgview1.Selected.Text,';',strlist,true);
    cbrd := TClipboard.Create;
    cbrd.SetTextBuf(PChar(trim(strlist.Strings[1])));
    freeandnil(strlist);
    freeandnil(cbrd);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.msgview1KeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then begin
    msgview1DblClick(Sender);
    Key := #0;
  end
end;

procedure TfmMain.SaveUserSelecion(basewrd : string; selwrd : string);
var dconf : AspellConfig; is_dict_error : AspellCanHaveError; spchk : AspellSpeller;
begin
  try
    if (not((trim(basewrd) = '') or (trim(selwrd)=''))) then begin
      dconf := new_aspell_config();
      aspell_config_replace(dconf,'lang',PChar(GetActiveDict));
      is_dict_error := new_aspell_speller(dconf);
      spchk := to_aspell_speller(is_dict_error);
      delete_aspell_config(dconf);
      aspell_speller_store_replacement(spchk,Pchar(basewrd),Length(basewrd),PChar(selwrd),Length(selwrd));
      aspell_speller_save_all_word_lists(spchk);
      delete_aspell_speller(spchk);
    end;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.btnSettingsClick(Sender: TObject);
begin
  try
    shellexecute(0,'open','http://code.google.com/p/spellcode','','',SW_SHOW);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.lstFilesKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then begin
    lstFilesDblClick(Sender);
    Key := #0;
  end;
end;

procedure TfmMain.btnCopyClick(Sender: TObject);
begin
  if (editor1.Enabled) then
    editor1.CopyToClipboard;
end;

procedure TfmMain.btnPasteClick(Sender: TObject);
begin
  if (editor1.Enabled) then
    editor1.PasteFromClipboard;
end;

procedure TfmMain.btnCutClick(Sender: TObject);
begin
  if (editor1.Enabled) then
    editor1.CutToClipboard;
end;

procedure TfmMain.btnUndoClick(Sender: TObject);
begin
  if (editor1.Enabled) then
    editor1.Undo;
end;

procedure TfmMain.btnRedoClick(Sender: TObject);
begin
  if (editor1.Enabled) then
    editor1.Redo;
end;

procedure TfmMain.MRUListInsertion(str : string; mru_file : string);
var mru : TStringList; pos : integer; str_ : string; proc_ : Boolean;
begin
  try
    if trim(str) = '' then
      exit;
    proc_ := false;
    mru := TStringList.Create;
    if fileexists(mru_file) then
      mru.LoadFromFile(mru_file);
    for pos := 0 to (mru.Count - 1) do begin
      if (lowercase(mru.Strings[pos]) = lowercase(str)) then begin
        str_ := mru.Strings[pos];
        mru.Delete(pos);
        mru.Insert(0,str_);
        proc_ := true;
        break;
      end;
    end;
    if not proc_ then begin
      mru.Insert(0,str);
      if mru.Count > 5 then
        mru.Delete(5);
    end;
    mru.SaveToFile(mru_file);
    freeandnil(mru);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;
end;

procedure TfmMain.FormShow(Sender: TObject);
var fext_list : TStringList;

  procedure ShowHelpMessage();
  begin
    MessageBox(handle,pchar('Command line parameters for spellcode:'#13#13'spellcode[.exe] [parameters]'#13#13'Valid Parameters:'#13#13'?'#9#9#9#9'Show help message'#13'<directory> <language index>'#9'Scan the specified directory'#13#13'For more information check http://code.google.com/p/spellcode'),pchar(Application.Title),MB_OK+MB_ICONINFORMATION)
  end;

begin
  try
    if (ParamCount > 0) then begin
      if ((trim(ParamStr(1)) = '?') or (trim(ParamStr(1)) = '/?')) then
        ShowHelpMessage
      else begin
        if ((ParamCount = 2) and (SysUtils.DirectoryExists(ParamStr(1))) and ((StrToInt(ParamStr(2))>0) and (StrToInt(ParamStr(2))<11))) then begin
          cmbLanguage.ItemIndex := (StrToInt(ParamStr(2)) - 1);
          Application.ProcessMessages;
          cmbLanguageChange(Sender);
          scan_dir := false;
		      status_panel.Panels[0].Text := 'Scanning the directory contents...';
          fext_list := TStringList.Create;
		      GetSupportFileExt(cmbLanguage.Items.Strings[cmbLanguage.ItemIndex], fext_list);
		      lstFiles.Items.Clear;
          BuildScanFileList(lstFiles.Items.AddChild(nil, ExtractFileName(ExcludeTrailingBackslash(ParamStr(1)))), file_list, ParamStr(1), fext_list);
		      freeandnil(fext_list);
          status_panel.Panels[0].Text := '';
          self.btnRescan.Enabled := true;
          lstFiles.Items.GetFirstNode.Expand(false);
          Clear1Click(sender);
          MRUListInsertion(ParamStr(1),root+'spellcode.mru');
          dirbrowse.Directory := ParamStr(1);
          Application.ProcessMessages;
          if not app_settings.dict_select_message then begin
            qhelp.ActivateHint(btnLanguage,'Before continue with the SpellCode make sure to select appropriate dictionary from here.','Select the dictionary from here...',6000);
            app_settings.dict_select_message := true;
          end;
        end else
          ShowHelpMessage;
      end;
    end else if (app_settings.show_startup_win) then
      startup_win.Enabled := true;
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.startup_winTimer(Sender: TObject);
var startwin : TuStartupWin; fext_list : TStringList;
begin
  try
    startup_win.Enabled := false;
    startwin := TuStartupWin.Create(self);
    if startwin.ShowModal = mrOK then begin
      app_settings.show_startup_win := not startwin.chkStartup.Checked;
      cmbLanguage.ItemIndex := startwin.cmbLangs.ItemIndex;
      cmbLanguageChange(Sender);
      scan_dir := false;
      status_panel.Panels[0].Text := 'Scanning the directory contents...';
      fext_list := TStringList.Create;
      GetSupportFileExt(cmbLanguage.Items.Strings[cmbLanguage.ItemIndex], fext_list);
      lstFiles.Items.Clear;
      BuildScanFileList(lstFiles.Items.AddChild(nil, ExtractFileName(ExcludeTrailingBackslash(startwin.tvDirs.Directory))), file_list, startwin.tvDirs.Directory, fext_list);
      freeandnil(fext_list);
      status_panel.Panels[0].Text := '';
      self.btnRescan.Enabled := true;
      lstFiles.Items.GetFirstNode.Expand(false);
      Clear1Click(sender);
      MRUListInsertion(startwin.tvDirs.Directory,root+'spellcode.mru');
      dirbrowse.Directory := startwin.tvDirs.Directory;
      if not app_settings.dict_select_message then begin
        qhelp.ActivateHint(btnLanguage,'Before continue with the SpellCode make sure to select appropriate dictionary from here.','Select the dictionary from here...',6000);
        app_settings.dict_select_message := true;
      end;
    end;
    freeandnil(startwin);
  except on E: Exception do
    MessageBox(handle,pchar(general_err_txt+E.Message),pchar(Application.Title + ' - Error'),MB_OK+MB_ICONERROR);
  end;  
end;

procedure TfmMain.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    113 : btnSaveClick(btnSave);
    114 : btnBrowseClick(btnBrowse);
    116 : btnScanClick(btnScan);
    118 : btnCheckSpellClick(btnCheckSpell);
  end;
end;

end.


