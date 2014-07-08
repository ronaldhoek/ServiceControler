unit ServiceStarterU;

interface

(*
  Based on the service control component made bij Masterijn (January 2004)
  
  http://programmersheaven.com/discussion/234511/stopping-and-starting-services-in-delphi-code
*)

uses
  WinSvc, Windows, Messages, SysUtils, Classes;

type
  TServiceState = (svsStopped, svsStarting, svsStopping, svsRunning,
    scvContinueing, svsPausing, svsPaused);

  TServiceStarter = class(TComponent)
  private
    FActive: Boolean;
    FHandle: THandle;
    FMachineName: string;
    FSCHandle: THandle;
    FServiceDisplayName: string;
    FServiceName: string;
    FState: TServiceState;
    FStateSet: Boolean;
    procedure CloseDependendServices(Handle: THandle);
    function GetHandle: THandle;
    function GetState: TServiceState;
    procedure SetActive(const Value: Boolean);
    procedure SetMachineName(const Value: string);
    procedure SetServiceDisplayName(const Value: string);
    procedure SetServiceName(const Value: string);
    procedure SetState(const Value: TServiceState);
    function StoreServiceDisplayName: Boolean;
  protected
    procedure CloseHandle;
    procedure CloseHandleSC;
    procedure HandleNeeded;
  public
    destructor Destroy; override;
    property Handle: THandle read GetHandle;
  published
    property MachineName: string read FMachineName write SetMachineName;
    property ServiceDisplayName: string read FServiceDisplayName write
        SetServiceDisplayName stored StoreServiceDisplayName;
    property ServiceName: string read FServiceName write SetServiceName;
    property State: TServiceState read GetState write SetState default svsStopped;
    // Must be last property!!!
    property Active: Boolean read FActive write SetActive default False;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Services', [TServiceStarter]);
end;

{ TServiceStarter }

procedure TServiceStarter.CloseDependendServices(Handle: THandle);
type
  TEnumServiceStatusArray = array [0 .. $FFFF] of TEnumServiceStatus;
  PEnumServiceStatusArray = ^TEnumServiceStatusArray;
var
  DependendHandle: THandle;
  I: Integer;
  ServicesCount: Cardinal;
  NeededBytes: Cardinal;
  DependendServices: PEnumServiceStatusArray;
//  ServiceStatus: TServiceStatus;
begin
  NeededBytes := 1 * SizeOf(TEnumServiceStatusArray);
  GetMem(DependendServices, NeededBytes);
  while not EnumDependentServices(Handle, SERVICE_ACTIVE, DependendServices^[0],
    NeededBytes, NeededBytes, ServicesCount) do
  begin
    if GetLastError = ERROR_MORE_DATA then
    begin
      FreeMem(DependendServices);
      GetMem(DependendServices, NeededBytes);
    end
    else
      RaiseLastOSError;
  end;
  for I := 0 to ServicesCount - 1 do
    with DependendServices^[I] do
    begin
      DependendHandle := OpenService(FSCHandle, lpServiceName,
        GENERIC_EXECUTE + SERVICE_QUERY_STATUS + SERVICE_ENUMERATE_DEPENDENTS);
      while not ControlService(DependendHandle, SERVICE_CONTROL_STOP,
        ServiceStatus) do
        case GetLastError of
          ERROR_DEPENDENT_SERVICES_RUNNING:
            CloseDependendServices(DependendHandle);
          ERROR_SERVICE_NOT_ACTIVE:
            Break;
          ERROR_SERVICE_CANNOT_ACCEPT_CTRL:
            ;
        else
          RaiseLastOSError;
        end;
      CloseServiceHandle(DependendHandle);
    end;
  FreeMem(DependendServices);
end;

procedure TServiceStarter.CloseHandle;
begin
  if FHandle <> 0 then
  begin
    CloseServiceHandle(FHandle);
    FHandle := 0;
  end;
end;

procedure TServiceStarter.CloseHandleSC;
begin
  if FSCHandle <> 0 then
  begin
    CloseServiceHandle(FSCHandle);
    FSCHandle := 0;
  end;
end;

destructor TServiceStarter.Destroy;
begin
  CloseHandle;
  CloseHandleSC;
  inherited;
end;

function TServiceStarter.GetHandle: THandle;
begin
  HandleNeeded;
  Result := FHandle;
end;

function TServiceStarter.GetState: TServiceState;
var
  ServiceStatus: TServiceStatus;
begin
  if FActive then
  begin
    if (FServiceName = '') and (FServiceDisplayName = '') then
    begin
      Result := svsStopped;
      Exit;
    end;
    HandleNeeded;
    if not QueryServiceStatus(FHandle, ServiceStatus) then
      RaiseLastOSError;
    Result := TServiceState(ServiceStatus.dwCurrentState - 1);
  end
  else
    Result := FState;
end;

procedure TServiceStarter.HandleNeeded;
var
  BuffSize: DWORD;
begin
  // Open service controler when needed
  if FSCHandle = 0 then
  begin
    FSCHandle := OpenSCManager(Pointer(FMachineName), nil, GENERIC_EXECUTE);
    if FSCHandle = 0 then
      RaiseLastOSError;
  end;

  // Open service when needed
  if FHandle = 0 then
  begin
    // Get service name based on service displayname (when available)
    if (Length(FServiceName) = 0) and (Length(FServiceDisplayName) > 0) then
    begin
      GetServiceKeyName(FSCHandle, PChar(FServiceDisplayName), nil, BuffSize);
      Inc(BuffSize); // Extra nul terminating char
      SetLength(FServiceName, BuffSize);
      if GetServiceKeyName(FSCHandle, PChar(FServiceDisplayName), PChar(FServiceName), BuffSize) then
        FServiceName := StrPas(PChar(FServiceName)) // Remove nul terminating char
      else begin
        FServiceName := '';
        RaiseLastOSError;
      end;
    end;

    FHandle := OpenService(FSCHandle, PChar(FServiceName),
      GENERIC_EXECUTE + SERVICE_QUERY_STATUS + SERVICE_ENUMERATE_DEPENDENTS);

    if FHandle = 0 then
      RaiseLastOSError;

    // Servicename is correct, look for displayname now (if not available yet)
    if Length(FServiceDisplayName) = 0 then
    begin
      GetServiceDisplayName(FSCHandle, PChar(FServiceName), nil, BuffSize);
      Inc(BuffSize); // Extra nul terminating char
      SetLength(FServiceDisplayName, BuffSize);
      if GetServiceDisplayName(FSCHandle, PChar(FServiceName), PChar(FServiceDisplayName), BuffSize) then
        FServiceDisplayName := StrPas(PChar(FServiceDisplayName)) // Remove nul terminating char
      else begin
        FServiceDisplayName := '';
        RaiseLastOSError;
      end;
    end;
  end;
end;

procedure TServiceStarter.SetActive(const Value: Boolean);
var
  CurState: TServiceState;
begin
  if FActive = Value then
    Exit;

  FActive := Value;
  if FActive then
  begin
    CurState := FState; // Save current internal state
    FState := GetState; // Request actual state of service

    // Internal state explicitly set when inactive?
    if FStateSet then
      SetState(CurState); // Then set status
  end else
  begin
    CloseHandle;
    CloseHandleSC;
  end;
end;

procedure TServiceStarter.SetMachineName(const Value: string);
begin
  if FMachineName = Value then
    Exit;

  CloseHandle;
  CloseHandleSC;
  FMachineName := Value;
end;

procedure TServiceStarter.SetServiceDisplayName(const Value: string);
begin
  if FServiceDisplayName = Value then
    Exit;
  CloseHandle;
  FServiceDisplayName := Value;
  // Reset servicename
  FServiceName := '';
end;

procedure TServiceStarter.SetServiceName(const Value: string);
begin
  if FServiceName = Value then
    Exit;
  CloseHandle;
  FServiceName := Value;
  // Reset service displayname
  FServiceDisplayName := '';
end;

procedure TServiceStarter.SetState(const Value: TServiceState);
const
  {
    SERVICE_CONTROL_STOP
    Requests the service to stop. The hService handle must have SERVICE_STOP access.
    SERVICE_CONTROL_PAUSE
    Requests the service to pause. The hService handle must have SERVICE_PAUSE_CONTINUE access.
    SERVICE_CONTROL_CONTINUE
    Requests the paused service to resume. The hService handle must have SERVICE_PAUSE_CONTINUE access.
    SERVICE_CONTROL_INTERROGATE
    Requests the service to update immediately its current status information to the service control manager. The hService handle must have SERVICE_INTERROGATE access.
    SERVICE_CONTROL_SHUTDOWN
  }

//  TServiceState = (svsStopped, svsStarting, svsStopping, svsRunning, scvContinueing, svsPausing, svsPaused);

  StateControlMap: array [TServiceState] of Integer = (SERVICE_CONTROL_STOP,
    SERVICE_CONTROL_CONTINUE, SERVICE_CONTROL_STOP, SERVICE_CONTROL_CONTINUE,
    SERVICE_CONTROL_CONTINUE, SERVICE_CONTROL_PAUSE, SERVICE_CONTROL_PAUSE);
var
  Error: Cardinal;
  StateSet: Boolean;
  Arg: PChar;
  ServiceStatus: TServiceStatus;
begin
  // State explicitly been set?
  if not Active then
    FStateSet := True;

  if FState = Value then
    Exit;
  FState := Value;

  if Active then
  begin
    HandleNeeded;
    Arg := nil;
    StateSet := False;
    // svsStopped, svsStarting, svsStopping, svsRunning, scvContinueing, svsPausing, svsPaused
    repeat
      if not ControlService(FHandle, StateControlMap[Value], ServiceStatus) then
      begin
        Error := GetLastError;
        case Error of
          ERROR_SERVICE_CANNOT_ACCEPT_CTRL:
            begin
              Sleep(10);
            end;
          ERROR_SERVICE_NOT_ACTIVE:
            if not(Value in [svsStopped, svsStopping]) then
            begin
              if not StartService(FHandle, 0, Arg) then
              begin
                Error := GetLastError;
                if Error <> ERROR_SERVICE_CANNOT_ACCEPT_CTRL then
                  RaiseLastOSError
                else
                  Sleep(10);
              end;
              StateSet := Value in [svsRunning, scvContinueing];
            end
            else
              StateSet := True;
          ERROR_DEPENDENT_SERVICES_RUNNING:
            CloseDependendServices(FHandle);
        else
          RaiseLastOSError;
        end;
      end
      else
        StateSet := True;
    until StateSet;
    FStateSet := False;
  end;
end;

function TServiceStarter.StoreServiceDisplayName: Boolean;
begin
  Result := (Length(FServiceDisplayName) > 0) and (Length(FServiceName) = 0);
end;

end.
