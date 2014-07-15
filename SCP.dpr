program SCP;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  WinApi.Windows,
  System.SysUtils,
  System.TypInfo,
  ServiceStarterU in 'ServiceStarterU.pas';

function HasParam(const sParam: string): Boolean;
var
  I: Integer;
begin
  for I := 1 to ParamCount do
    if SameText(ParamStr(I), sParam) then
      Exit(True);
  Result := False;
end;

const
  WaitTimeOut = 5000; // 5 seconds
var
  NewState: TServiceState;
  Start: Cardinal;
begin
  try
    with TServiceStarter.Create(nil) do
    try
      MachineName := '';
      // parameter to specify wether the name specified is the service
      // displayname or the actual servicename.
      if HasParam('/d') then
        ServiceDisplayName := ParamStr(1)
      else
        ServiceName := ParamStr(1);
      Active := True;
      Writeln('Service: ', ServiceName, ' / ', ServiceDisplayName);
      Writeln('State: ', GetEnumName(TypeInfo(TServiceState), Integer(State)));
      // Set a new state?
      NewState := State; // Get current first
      if HasParam('/start') then
        NewState := svsRunning
      else if HasParam('/stop') then
        NewState := svsStopped
      else if HasParam('/pause') then
        NewState := svsPaused;
      if NewState <> State then
      begin
        Writeln('Setting new state: ', GetEnumName(TypeInfo(TServiceState), Integer(NewState)));
        Start := GetTickCount;
        State := NewState;
        while (State <> NewState) and (GetTickCount < Start + WaitTimeOut) do
          Sleep(100);
        Writeln('New state: ', GetEnumName(TypeInfo(TServiceState), Integer(State)));
      end;
    finally
      Free;
    end;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.

