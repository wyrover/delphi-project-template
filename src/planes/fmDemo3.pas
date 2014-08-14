unit fmDemo3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, VirtualTrees, ExtCtrls, OleCtrls, SHDocVw, cmpExWebBrowser,
  ComObj, ActiveX, MSHTML, SHDocVw_EWB, EwbCore, EmbeddedWB;

type
  TSendLinkEvent = procedure(Sender: TObject; const ALink, AText: string) of object;
  TAiGuangJieThread = class(TThread)
  private
    Fsmp:        string;
    FLink:       string;
    FText:       string;
    FGroupCount: integer;
    FOnSendLink: TSendLinkEvent;
    procedure SendLink;
  protected
    procedure Execute; override;
  published
    property OnSendLink: TSendLinkEvent read FOnSendLink write FOnSendLink;
  end;


  TZhuanjiThread = class(TThread)
  private
    FLink:       string;
    FText:       string;
    FieApp:      Variant;
    FOnSendLink: TSendLinkEvent;
    procedure SendLink;
    function GetHtml(const url: string): string;
    function  ProcessZhuanji(const url: string): Boolean;
  protected
    procedure Execute; override;
  published
    property  OnSendLink: TSendLinkEvent read FOnSendLink write FOnSendLink;
  end;

  TIEBeforeNavigate2Event = procedure(Sender: TObject; const pDisp: IDispatch;
    var URL: OleVariant; var Flags: OleVariant; var TargetFrameName: OleVariant;
    var PostData: OleVariant; var Headers: OleVariant; var Cancel: WordBool) of object;

  TAiGuangJieAutoIEThread = class(TThread, IUnknown, IDispatch)
  private
    FLink:            string;
    FText:            string;
    FieApp:           IWebBrowser2;
    FConnected:       Boolean;
    FSinkIID:         TGuid;
    FCPContainer:     IConnectionPointContainer;
    FCP:              IConnectionPoint;
    FCookie:          Integer;
    FBeforeNavigate2: TIEBeforeNavigate2Event;
    FOnSendLink:      TSendLinkEvent;
    procedure SendLink;
    function  GetHtml(const url: string): string;
    function  InsertZhuanjiTable(const url, caption: string): Boolean;
    function  ProcessZhuanji(const url: string): Boolean;
    procedure ConnectTo(Source: IWebBrowser2);
    procedure Disconnect;
  protected
    procedure Execute; override;

    // Protected declaratios for IUnknown
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
    // Protected declaratios for IDispatch
    function GetIDsOfNames(const IID: TGUID; Names: Pointer; NameCount, LocaleID:
      Integer; DispIDs: Pointer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    // 处理IE事件
    procedure DoBeforeNavigate2(const pDisp: IDispatch; var URL: OleVariant;
      var Flags: OleVariant; var TargetFrameName: OleVariant; var PostData: OleVariant;
      var Headers: OleVariant; var Cancel: WordBool); safecall;
  published
    property OnSendLink: TSendLinkEvent read FOnSendLink write FOnSendLink;
    property OnBeforeNavigate2: TIEBeforeNavigate2Event read FBeforeNavigate2 write FBeforeNavigate2;
  end;






  PTreeData = ^TTreeData;
  TTreeData = record
    FLink: string;
    FText: string;
  end;




  TDemo3Form = class(TForm)
    vTree: TVirtualStringTree;
    spl1: TSplitter;
    btn1: TButton;
    btn2: TButton;
    btn3: TButton;
    btn4: TButton;
    EmbeddedWB1: TEmbeddedWB;
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure btn3Click(Sender: TObject);
    procedure btn4Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure vTreeFreeNode(Sender: TBaseVirtualTree; Node: PVirtualNode);
    procedure vTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode; Column:
        TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure wb1BeforeNavigate2(ASender: TObject; const pDisp: IDispatch; var URL,
        Flags, TargetFrameName, PostData, Headers: OLEVariant; var Cancel:
        WordBool);
  private
    procedure InitTreeView;
    procedure SendLink(Sender: TObject; const ALink, AText: string);
    procedure SetWebrowserURL(Sender: TObject; const ALink, AText: string);
    procedure DocumentComplete(Sender: TObject; doc: IHTMLDocument);
    procedure BeforeNavigate2(Sender: TObject; const pDisp: IDispatch;
    var URL: OleVariant; var Flags: OleVariant; var TargetFrameName: OleVariant;
    var PostData: OleVariant; var Headers: OleVariant; var Cancel: WordBool);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Demo3Form: TDemo3Form;

implementation

uses
  RegularExpressions, TraceTool, appUtils, roShell, roFiles, IdURI, SQLiteTable3,
  roStrings, StrUtils, roAutoIE;

{$R *.dfm}

{ TYonghengThread }

procedure TAiGuangJieThread.Execute;
var
  input: string;
  Regex: IRegex;
  match: IMatch;
  dir:   string;
  page:  Integer;
  idurl: TIdURI;
  db:    TSQLiteDatabase;
  tb:    TSQLiteTable;
  sql:   string;
begin
  inherited;

  dir := NormalDir(GetAppPath + 'aiguangjie');
  CreateDir(dir);

  page := 1;

  Regex := TRegex.Create('<div\s+class="block"\s+data-spm="(.*?)"[^\b]+?<a\s+class="c3\s+yh"\s+href="(.*?)".*>(.*?)</a>', [roIgnoreCase, roMultiline]);

  try
    while page > 0 do
    begin
      input := GetHtml('http://love.taobao.com/album/index.htm?spm=1001.1000484.0.462.UKACAP&order=1&cid=0&q=&tab=0&tagid=1&page=' + IntToStr(page));
      match := Regex.Match(input);

      if match.Success then
        Inc(page)
      else
        page := 0;

      while match.Success do
      begin
        self.Fsmp := match.Groups[1].Value;
        Self.FLink := ReplaceStr(match.Groups[2].value, 'detail.htm?', 'detail.htm?spm=' + match.Groups[1].Value + '&');
        Self.FText := match.Groups[3].value;
        Synchronize(self.SendLink);

        WriteTextToFile(GetAppPath + 'link.txt', Self.FLink, true);

        sql := 'INSERT INTO zhuanji(url, caption) VALUES ("' + self.Flink + '","' + Self.FText + '")';
        db := GetAiGuangJieDb;
        try
          tb := db.GetTable('SELECT * FROM zhuanji where url = "' + Self.FLink + '"');
          if tb.Count = 0 then db.ExecSQL(sql);
          tb.Free;
        finally
          db.Free;
        end;



        //CreateDir(NormalDir(dir+ Self.FText));

        {idurl := TIdURI.Create(Self.FLink);
        if not FileExists(dir + idurl.Document) then
          DownloadToFile(Self.FLink, dir + idurl.Document);
        idurl.Free;}
        match := match.NextMatch;
      end;
    end;
  finally
    Regex := nil;
  end;


end;

procedure TAiGuangJieThread.SendLink;
begin
  if Assigned(FOnSendLink) then
    FOnSendLink(Self, FLink, FText);
end;


procedure TDemo3Form.BeforeNavigate2(Sender: TObject; const pDisp: IDispatch;
  var URL, Flags, TargetFrameName, PostData, Headers: OleVariant;
  var Cancel: WordBool);
begin
  if ContainsStr(URL, 'detail.htm') then
    Cancel := True;

  //TTrace.Debug.Send('url', URL);

end;

procedure TDemo3Form.btn1Click(Sender: TObject);
begin
  with TAiGuangJieAutoIEThread.Create(false) do
  begin
    OnSendLink := self.SendLink;
    OnBeforeNavigate2 := Self.BeforeNavigate2;
    Resume;
  end;
end;

procedure TDemo3Form.btn2Click(Sender: TObject);
begin
  with TZhuanjiThread.Create(false) do
  begin
    OnSendLink := self.SetWebrowserURL;
    Resume;
  end;
end;

procedure TDemo3Form.btn3Click(Sender: TObject);
var
  ieApp: Variant;
begin

  ieApp := CreateOleObject('InternetExplorer.Application');
  ieApp.visible := True;
  ieApp.AddressBar := True;
  ieApp.menubar := True;
  ieApp.ToolBar := True;
  ieApp.StatusBar := True;
  ieApp.width := 800;
  ieApp.height := 600;
  ieApp.resizable := True;
  ieApp.Navigate('http://love.taobao.com/album/index.htm');


  while ieApp.ReadyState = 4 do Break;

  ShowMessage('加载完毕');

  WriteTextToFile(GetAppPath + 'out.htm', ieApp.Document.body.outerHTML, False);



end;

procedure TDemo3Form.btn4Click(Sender: TObject);
var
  autoIE: TAutoIE;
begin
  autoIE := TAutoIE.Create;
  try
    autoIE.OnDocumentComplete := self.DocumentComplete;
    autoIE.Navigate('http://www.baidu.com');
  finally
    autoIE.Free;
  end;
end;

procedure TDemo3Form.DocumentComplete(Sender: TObject; doc: IHTMLDocument);
begin
  ShowMessage((doc as IHTMLDocument2).body.outerHTML);
end;

procedure TDemo3Form.FormDestroy(Sender: TObject);
begin
  TTrace.Debug.EnterMethod('TDemo3Form.FormDestroy');
  TTrace.Debug.ExitMethod('TDemo3Form.FormDestroy');
end;

procedure TDemo3Form.FormCreate(Sender: TObject);
begin


  GetAiGuangjieDb;
  InitTreeView;
  
end;

procedure TDemo3Form.InitTreeView;
var
  rootNode: PVirtualNode;
  childNode: PVirtualNode;
  Data: PTreeData;
begin
  vTree.NodeDataSize := SizeOf(TTreeData);
  //vTree.Images := ImageListMain;

  //XML2Tree(vTree, ChangeFileExt(ParamStr(0),'.XML'));

  vTree.TreeOptions.PaintOptions := vTree.TreeOptions.PaintOptions -[toShowRoot];
  //vTree.FullExpand;

end;

procedure TDemo3Form.SendLink(Sender: TObject; const ALink, AText: string);
var
  Data: PTreeData;
  tn: PVirtualNode;
begin
  tn := vTree.AddChild(nil);
  Data := vTree.GetNodeData(tn);
  Data^.FLink := alink;
  Data^.FText := aText;
end;

procedure TDemo3Form.SetWebrowserURL(Sender: TObject; const ALink,
  AText: string);
begin

  //wb1.Navigate(alink);
end;

procedure TDemo3Form.vTreeFreeNode(Sender: TBaseVirtualTree; Node:
    PVirtualNode);
var
  Data: PTreeData;
begin
  Data := Sender.GetNodeData(Node);
  if Assigned(Data) then
    Finalize(Data^);
end;

procedure TDemo3Form.vTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
    Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
var
  Data: PTreeData;
begin
  Data:=Sender.GetNodeData(Node);
  CellText := Data^.FLink;
end;

procedure TDemo3Form.wb1BeforeNavigate2(ASender: TObject; const pDisp:
    IDispatch; var URL, Flags, TargetFrameName, PostData, Headers: OLEVariant;
    var Cancel: WordBool);
begin

end;

{ TZhuanjiThread }

procedure TZhuanjiThread.Execute;
var
  db:    TSQLiteDatabase;
  tb:    TSQLiteTable;
  sql:   string;
begin
  inherited;
  CoInitialize(nil);
  sql := 'SELECT * FROM zhuanji';
  db := GetAiGuangJieDb;
  try
    tb := db.GetTable(sql);
    if tb.Count > 0 then
    begin
      tb.MoveFirst;
      while not tb.EOF do
      begin
        ProcessZhuanji(tb.FieldAsString(tb.FieldIndex['url']) + '&page=');
        tb.Next;
      end;
    end;
    tb.Free;
  finally
    db.Free;
  end;
  CoUninitialize;
end;

function TZhuanjiThread.GetHtml(const url: string): string;
begin
  if VarIsEmpty(FieApp) then
  begin
    FieApp := CreateOleObject('InternetExplorer.Application');
    FieApp.visible := True;
    FieApp.AddressBar := True;
    FieApp.menubar := True;
    FieApp.ToolBar := True;
    FieApp.StatusBar := True;
    FieApp.width := 800;
    FieApp.height := 600;
    FieApp.resizable := True;
  end;

  FieApp.Navigate(url);
  while (FieApp.busy) or (FieApp.ReadyState <>READYSTATE_COMPLETE) do
    Application.ProcessMessages;
  Result := FieApp.Document.body.outerHTML;
end;

function TZhuanjiThread.ProcessZhuanji(const url: string): Boolean;
var
  input: string;
  Regex: IRegex;
  match: IMatch;
  Group: IGroup;
  dir:   string;
  page:  Integer;
  idurl: TIdURI;
  db:    TSQLiteDatabase;
  tb:    TSQLiteTable;
  sql:   string;
  I:     Integer;
begin
  dir := NormalDir(GetAppPath + 'aiguangjie_image');
  CreateDir(dir);

  page := 1;

  //Regex := TRegex.Create('<a\s+class="J_IAddAlbumLink\s+l-j-add".*?data-infoid="(.*?)"\s+data-infouid="(.*?)"\s+data-id="(.*?)"\s+data-type="(.*?)"\s+data-img="(.*?)".*?data-aid="(.*?)"\s+>', [roIgnoreCase, roMultiline]);

  Regex := TRegex.Create('<a\s+class="J_IAddAlbumLink\s+l-j-add"[^\b]+?data-type="(.*?)"\s+data-aid="(.*?)"\s+data-id="(.*?)"\s+data-img="(.*?)"\s+data-infouid="(.*?)"\s+data-infoid="(.*?)"', [roIgnoreCase, roMultiline]);

  try
    while page > 0 do
    begin
      input := Self.GetHtml(url + IntToStr(page));
      //WriteTextToFile(GetAppPath + 'input.htm', input, False);

      match := Regex.Match(input);

      if match.Success then
      begin
        Inc(page);
      end else begin
        page := 0;
      end;




      while match.Success do
      begin
        Inc(I);
        Self.FLink := ReplaceStr(match.Groups[4].value, '_120x120.jpg', '');
        //Self.FText := match.Groups[2].value;
        //Self.FGroupCount := match.Groups.Count;
        //Synchronize(self.SendLink);

        //CreateDir(NormalDir(dir+ Self.FText));

        idurl := TIdURI.Create(Self.FLink);
        if not FileExists(dir + idurl.Document) then
          DownloadToFile(Self.FLink, dir + idurl.Document);
        idurl.Free;
        match := match.NextMatch;
        if I < 40 then page := 0;        
      end;
    end;
  finally
    Regex := nil;
  end;


end;

procedure TZhuanjiThread.SendLink;
begin
  if Assigned(FOnSendLink) then
    FOnSendLink(Self, FLink, FText);
end;

{ TAiGuangJieAutoIEThread }

procedure TAiGuangJieAutoIEThread.ConnectTo(Source: IWebBrowser2);
var
  pvCPC: IConnectionPointContainer;
begin
  // Disconnect from any currently connected event sink
  Disconnect;
  // Query for the connection point container and desired connection point.
  // On success, sink the connection point
  OleCheck(Source.QueryInterface(IConnectionPointContainer, pvCPC));
  OleCheck(pvCPC.FindConnectionPoint(FSinkIID, FCP));
  OleCheck(FCP.Advise(Self, FCookie));
  // Update internal state variables
  FieApp := Source;
  // We are in a connected state
  FConnected := True;
  // Release the temp interface
  pvCPC := nil;
end;

procedure TAiGuangJieAutoIEThread.Disconnect;
begin
  // Do we have the IWebBrowser2 interface?
  if Assigned(FieApp) then
  begin
    try
      if Assigned(FCP) then
        OleCheck(FCP.Unadvise(FCookie));
        // Release the interfaces
      FCP := nil;
    except
      Pointer(FCP) := nil;
      Pointer(FieApp) := nil;
    end;
  end;

  // Disconnected state
  FConnected := False;
end;

procedure TAiGuangJieAutoIEThread.DoBeforeNavigate2(const pDisp: IDispatch;
  var URL, Flags, TargetFrameName, PostData, Headers: OleVariant;
  var Cancel: WordBool);
begin
  if Assigned(FBeforeNavigate2) then
    FBeforeNavigate2(Self, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel);
end;

procedure TAiGuangJieAutoIEThread.Execute;
var
  input: string;
  Regex: IRegex;
  match: IMatch;
  dir:   string;
  page:  Integer;
  idurl: TIdURI;

begin
  inherited;
  CoInitialize(nil);

  dir := NormalDir(GetAppPath + 'aiguangjie');
  CreateDir(dir);

  page := 1;

  Regex := TRegex.Create('<a\s+class="c3\s+yh"\s+href="(.*?)".*>(.*?)</a>', [roIgnoreCase, roMultiline]);

  try
    while page > 0 do
    begin
      input := Self.GetHtml('http://love.taobao.com/album/index.htm?page=' + IntToStr(page));
      match := Regex.Match(input);

      if match.Success then
      begin
        Inc(page)
      end else begin
        page := 0;
      end;

      {while match.Success do
      begin
        Self.FLink := match.Groups[1].value;
        Self.FText := match.Groups[2].value;
        Synchronize(self.SendLink);
        Self.InsertZhuanjiTable(match.Groups[1].value, match.Groups[2].value);

        self.ProcessZhuanji('http://love.taobao.com/album/' + Self.FLink + '&page=');


        //CreateDir(NormalDir(dir+ Self.FText));

        {idurl := TIdURI.Create(Self.FLink);
        if not FileExists(dir + idurl.Document) then
          DownloadToFile(Self.FLink, dir + idurl.Document);
        idurl.Free;
        match := match.NextMatch;
      end;}
    end;
  finally
    Regex := nil;
  end;
  CoUninitialize;
end;

function TAiGuangJieAutoIEThread.GetHtml(const url: string): string;
var
  IE:                OleVariant;
  doc3:              IHTMLDocument3;
  elementCollection: IHTMLElementCollection;
  element:           IHTMLElement4;
  I:                 Integer;
  a:                 OleVariant;
  flags, TargetFrameName, PostData, Headers: OleVariant;
begin
  if not Assigned(FieApp) then
  begin
    FSinkIID := DWebBrowserEvents2;
    IE := CreateOleObject('InternetExplorer.Application');
    if (IDispatch(IE).QueryInterface(IWebBrowser2, FieApp) = S_OK) then
    begin

      FieApp.visible := True;
      FieApp.AddressBar := True;
      FieApp.menubar := True;
      //FieApp.ToolBar := True;
      FieApp.StatusBar := True;
      FieApp.width := 800;
      FieApp.height := 600;
      FieApp.resizable := True;

      Self.Disconnect;

      OleCheck(FieApp.QueryInterface(IConnectionPointContainer, FCPContainer));
      OleCheck(FCPContainer.FindConnectionPoint(FSinkIID, FCP));
      OleCheck(FCP.Advise(Self, FCookie));

      // We are in a connected state
      FConnected := True;
      // Release the temp interface
      FCPContainer := nil;
    end;
  end;



  FieApp.Navigate(url, flags,TargetFrameName,PostData,Headers);
  while (FieApp.busy) or (FieApp.ReadyState <>READYSTATE_COMPLETE) do
    Application.ProcessMessages;

  Self.ConnectTo(FieApp);

  doc3 := FieApp.Document as IHTMLDocument3;
  elementCollection := doc3.getElementsByTagName('a');

  for I := 0 to Pred(elementCollection.length) do
  begin
    element := elementCollection.item(I, EmptyParam) as IHTMLElement4;
    if (element.getAttributeNode('class') <> nil) and (element.getAttributeNode('class').nodeValue = 'c3 yh') then
    begin
      element.getAttributeNode('target').nodeValue := '_self';
      {TTrace.Debug.Send('element', element.getAttributeNode('href').nodeValue);
      TTrace.Debug.Send('Text', (element as IHTMLElement).innerText);}
      Self.InsertZhuanjiTable(element.getAttributeNode('href').nodeValue, (element as IHTMLElement).innerText);
      a := elementCollection.item(I, EmptyParam);
      a.click;
    end;
  end;


  Result := (doc3 as IHTMLDocument2).body.outerHTML;
end;

function TAiGuangJieAutoIEThread.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  // Not implemented
  result := E_NOTIMPL;
end;

function TAiGuangJieAutoIEThread.GetTypeInfo(Index, LocaleID: Integer;
  out TypeInfo): HResult;
begin
  // Clear the result interface
  Pointer(TypeInfo) := nil;
  // No type info for our interface
  result := E_NOTIMPL;
end;

function TAiGuangJieAutoIEThread.GetTypeInfoCount(out Count: Integer): HResult;
begin
  // Zero type info counts
  Count := 0;
  // Return success
  result := S_OK;
end;

function TAiGuangJieAutoIEThread.InsertZhuanjiTable(const url,
  caption: string): Boolean;
var
  db:    TSQLiteDatabase;
  tb:    TSQLiteTable;
  sql:   string;
begin
  sql := 'INSERT INTO zhuanji(url, caption) VALUES ("' + url + '","' + caption + '")';
  db := GetAiGuangJieDb;
  try
    tb := db.GetTable('SELECT * FROM zhuanji where url = "' + url + '"');
    if tb.Count = 0 then db.ExecSQL(sql);
    tb.Free;
  finally
    db.Free;
  end;
end;

function TAiGuangJieAutoIEThread.Invoke(DispID: Integer; const IID: TGUID;
  LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo,
  ArgErr: Pointer): HResult;
var pdpParams: PDispParams;
  lpDispIDs: array[0..63] of TDispID;
  dwCount: Integer;
begin

  // Get the parameters
  pdpParams := @Params;

  // Events can only be called with method dispatch, not property get/set
  if ((Flags and DISPATCH_METHOD) > 0) then
  begin
     // Clear DispID list
    ZeroMemory(@lpDispIDs, SizeOf(lpDispIDs));
     // Build dispatch ID list to handle named args
    if (pdpParams^.cArgs > 0) then
    begin
        // Reverse the order of the params because they are backwards
      for dwCount := 0 to Pred(pdpParams^.cArgs) do lpDispIDs[dwCount] := Pred(pdpParams^.cArgs) - dwCount;
        // Handle named arguments
      if (pdpParams^.cNamedArgs > 0) then
      begin
        for dwCount := 0 to Pred(pdpParams^.cNamedArgs) do
          lpDispIDs[pdpParams^.rgdispidNamedArgs^[dwCount]] := dwCount;
      end;
    end;
     // Unless the event falls into the "else" clause of the case statement the result is S_OK
    result := S_OK;
     // Handle the event
    case DispID of
      {102: DoStatusTextChange(pdpParams^.rgvarg^[lpDispIds[0]].bstrval);
      104: DoDownloadComplete;
      105: DoCommandStateChange(pdpParams^.rgvarg^[lpDispIds[0]].lval,
          pdpParams^.rgvarg^[lpDispIds[1]].vbool);
      106: DoDownloadBegin;
      108: DoProgressChange(pdpParams^.rgvarg^[lpDispIds[0]].lval,
          pdpParams^.rgvarg^[lpDispIds[1]].lval);
      112: DoPropertyChange(pdpParams^.rgvarg^[lpDispIds[0]].bstrval);
      113: DoTitleChange(pdpParams^.rgvarg^[lpDispIds[0]].bstrval);}
      250: DoBeforeNavigate2(IDispatch(pdpParams^.rgvarg^[lpDispIds[0]].dispval),
          POleVariant(pdpParams^.rgvarg^[lpDispIds[1]].pvarval)^,
          POleVariant(pdpParams^.rgvarg^[lpDispIds[2]].pvarval)^,
          POleVariant(pdpParams^.rgvarg^[lpDispIds[3]].pvarval)^,
          POleVariant(pdpParams^.rgvarg^[lpDispIds[4]].pvarval)^,
          POleVariant(pdpParams^.rgvarg^[lpDispIds[5]].pvarval)^,
          pdpParams^.rgvarg^[lpDispIds[6]].pbool^);
      {251: DoNewWindow2(IDispatch(pdpParams^.rgvarg^[lpDispIds[0]].pdispval^),
          pdpParams^.rgvarg^[lpDispIds[1]].pbool^);
      252: DoNavigateComplete2(IDispatch(pdpParams^.rgvarg^[lpDispIds[0]].dispval),
          POleVariant(pdpParams^.rgvarg^[lpDispIds[1]].pvarval)^);
      253:
        begin
           // Special case handler. When Quit is called, IE is going away so we might
           // as well unbind from the interface by calling disconnect.
          DoOnQuit;
           //  Call disconnect
          Disconnect;
        end;
      254: DoOnVisible(pdpParams^.rgvarg^[lpDispIds[0]].vbool);
      255: DoOnToolBar(pdpParams^.rgvarg^[lpDispIds[0]].vbool);
      256: DoOnMenuBar(pdpParams^.rgvarg^[lpDispIds[0]].vbool);
      257: DoOnStatusBar(pdpParams^.rgvarg^[lpDispIds[0]].vbool);
      258: DoOnFullScreen(pdpParams^.rgvarg^[lpDispIds[0]].vbool);
      259: DoDocumentComplete(IDispatch(pdpParams^.rgvarg^[lpDispIds[0]].dispval),
          POleVariant(pdpParams^.rgvarg^[lpDispIds[1]].pvarval)^);
      260: DoOnTheaterMode(pdpParams^.rgvarg^[lpDispIds[0]].vbool);}
    else
        // Have to idea of what event they are calling
      result := DISP_E_MEMBERNOTFOUND;
    end;
  end
  else
     // Called with wrong flags
    result := DISP_E_MEMBERNOTFOUND;
end;

function TAiGuangJieAutoIEThread.ProcessZhuanji(const url: string): Boolean;
var
  input: string;
  Regex: IRegex;
  match: IMatch;
  Group: IGroup;
  dir:   string;
  page:  Integer;
  idurl: TIdURI;
  db:    TSQLiteDatabase;
  tb:    TSQLiteTable;
  sql:   string;
begin
  dir := NormalDir(GetAppPath + 'aiguangjie_image');
  CreateDir(dir);

  page := 1;

  Regex := TRegex.Create('<a\s+class="J_IAddAlbumLink\s+l-j-add".*?data-infoid="(.*?)"\s+data-infouid="(.*?)"\s+data-id="(.*?)"\s+data-type="(.*?)"\s+data-img="(.*?)".*?data-aid="(.*?)"\s+>', [roIgnoreCase, roMultiline]);

  try
    while page > 0 do
    begin
      input := Self.GetHtml(url + IntToStr(page));
      WriteTextToFile(GetAppPath + 'input.htm', input, false);

      match := Regex.Match(input);

      if match.Success then
        Inc(page)
      else
        page := 0;

      while match.Success do
      begin
        Self.FLink := ReplaceStr(match.Groups[5].value, '_120x120.jpg', '');
        //Self.FText := match.Groups[2].value;
        //Self.FGroupCount := match.Groups.Count;
        //Synchronize(self.SendLink);

        //CreateDir(NormalDir(dir+ Self.FText));

        idurl := TIdURI.Create(Self.FLink);
        if not FileExists(dir + idurl.Document) then
          DownloadToFile(Self.FLink, dir + idurl.Document);
        idurl.Free;
        match := match.NextMatch;
      end;
    end;
  finally
    Regex := nil;
  end;


end;

function TAiGuangJieAutoIEThread.QueryInterface(const IID: TGUID;
  out Obj): HResult;
begin
  // Clear interface pointer
  Pointer(Obj) := nil;

  // Attempt to get the requested interface
  if (GetInterface(IID, Obj)) then
     // Success
    result := S_OK
  // Check to see if the guid requested is for the event
  else if (IsEqualIID(IID, FSinkIID)) then
  begin
     // Event is dispatch based, so get dispatch interface (closest we can come)
    if (GetInterface(IDispatch, Obj)) then
        // Success
      result := S_OK
    else
        // Failure
      result := E_NOINTERFACE;
  end
  else
     // Failure
    result := E_NOINTERFACE;
end;

procedure TAiGuangJieAutoIEThread.SendLink;
begin

end;

function TAiGuangJieAutoIEThread._AddRef: Integer;
begin
  // No more than 2 counts
  result := 2;
end;

function TAiGuangJieAutoIEThread._Release: Integer;
begin
  // Always maintain 1 ref count (component holds the ref count)
  result := 1;
end;

initialization
  CoInitialize(nil);
finalization
  CoUninitialize;

end.
