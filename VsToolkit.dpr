(* Volume Shadow Toolkit
   =====================

   Converted from the C++ program "vscsc" (volume shadow copy simple client)
   refer to: http://sourceforge.net/projects/vscsc/
   also refer to: WindowsSDK v7.0: Samples\winbase\vss\vshadow\

   © Dr. J. Rathlev, D-24222 Schwentinental (info(a)rathlev-home.de)

   The contents of this file may be used under the terms of the
   Mozilla Public License ("MPL") or
   GNU Lesser General Public License Version 2 or later (the "LGPL")

   Software distributed under this License is distributed on an "AS IS" basis,
   WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for
   the specific language governing rights and limitations under the License.

   Vers. 1.0: October 2014
         2.0: March 2015
         2.1: January 2016
         2.2: May 2017
  *)

program VsToolkit;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  System.Classes,
  Winapi.Windows,
  Winapi.ActiveX,
  WinApiUtils,
  VssApi,
  VssUtils;

type
  TVssAction = (vaBackup,vaWriterStatus,vaWriterMetadata,vaWriterDetailedMetadata,
                vaQuery,vaQuerySet,vaQueryID,vaDeleteAll,vaDeleteSet,vaDeleteID,
                vaHelp);

const
  WelcomeMsg = 'VsToolkit 2.3.2 (%s) - Volume Shadow Toolkit for Windows 7/8/10'+sLineBreak+
    'This is a modified Delphi version of the Volume Shadow Copy Simple Client'+sLineBreak+
	  'by Microsoft (VSHADOW.EXE), originally bundled with the Microsoft SDK 7'+sLineBreak+
	  'for Windows Vista and Windows 7/8/10.'+sLineBreak;

{ ---------------------------------------------------------------- }
// Routinen zur Auswertung einer Befehlszeile
// prüfe, ob die ersten Zeichen einer Option mit dem Parameter übereinstimmen
function CompareOption (const Param,Option : string) : boolean;
begin
  Result:=AnsiLowercase(Param)=copy(Option,1,length(Param));
  end;

// Option vom Typ option:value einlesen
function ReadOptionValue (var Param : string; const Option : string) : boolean;
var
  i : integer;
begin
  Result:=false;
  i:=AnsiPos('=',Param);
  if i=0 then exit;
  if not CompareOption(copy(Param,1,i-1),Option) then exit;
  Delete(Param,1,i); Result:=true;
  end;


procedure ShowUsage;
begin
  Writeln('Usage:');
  Writeln('   VsToolkit [optional flags] [commands]');
  Writeln;
  Writeln('List of optional flags:');
  Writeln('  -?                 - Displays the usage screen');
  Writeln('  -wi={Writer Name}  - Verify that a writer/component is included');
  Writeln('  -wx={Writer Name}  - Exclude a writer/component from set creation or restore');
  Writeln('  -exec={command}    - Custom command executed after shadow creation.');
  Writeln('                       The command is invoked and the name of the new temporary');
  Writeln('                       shadow volume is passed as first parameter.');
  Writeln('  -wait              - Wait before program termination');
  Writeln;
  Writeln('List of commands:');
  Writeln('  {one volume}       - Creates a temporary shadow set on this volume');
  Writeln('                       (JUST ONE VOLUME AT A TIME!!!)');
  Writeln('  -ws                - List writer status');
  Writeln('  -wm                - List writer summary metadata');
  Writeln('  -wm2               - List writer detailed metadata');
//  Writeln('  -wm3               - List writer detailed metadata in raw XML format');
  Writeln('  -q                 - List all shadow copies in the system');
  Writeln('  -qx={SnapSetID}    - List all shadow copies in this set');
  Writeln('  -qs={SnapID}       - List the shadow copy with the given ID');
  Writeln('  -da                - Deletes all shadow copies in the system');
  Writeln('  -dx={SnapSetID}    - Deletes all shadow copies in this set');
  Writeln('  -ds={SnapID}       - Deletes this shadow copy');
  Writeln('  -log={filename}    - Write to log file');
  Writeln;
  end;

function UpdateAction (NewAction : TVssAction; var Action : TVssAction) : boolean;
begin
  if Action=vaBackup then begin
    Action:=NewAction; Result:=true;
    end
  else Result:=false;
  end;

var
  s,LogName,
  ExecCommand        : string;
  VolumeShadowCopy   : TVolumeShadowCopy;
  VolList,
  ExcludedWriterList,
  IncludedWriterList : TStringList;
  dwAttributes,
  i                  : integer;
  va                 : TVssAction;
  wait,ok            : boolean;
  SnapId,guid        : TGuid;
  hr                 : HResult;
begin
  CoInitialize(nil);
  Writeln;
  if Is64BitApp then s:='64-bit' else s:='32-bit';
  Writeln(Format(WelcomeMsg,[s]));
  Writeln;
  if ParamCount>0 then begin
    va:=vaBackup; Wait:=false; LogName:='';
    VolList:=TStringList.Create;
    ExcludedWriterList:=TStringList.Create;;
    IncludedWriterList:=TStringList.Create;
    execCommand:='';
    for i:=1 to ParamCount do begin
      s:=ParamStr(i);
      if length(s)>0 then begin
        if (s[1]='/') or (s[1]='-') then begin
          delete (s,1,1);
          if CompareOption(s,'?') or CompareOption(s,'help') then va:=vaHelp
          else if ReadOptionValue(s,'wi') then IncludedWriterList.Add(s)
          else if ReadOptionValue(s,'wx') then ExcludedWriterList.Add(s)
          else if ReadOptionValue(s,'exec') then ExecCommand:=s
          else if CompareOption(s,'wait') then Wait:=true
          else if CompareOption(s,'ws') then UpdateAction(vaWriterStatus,va)
          else if CompareOption(s,'wm') then UpdateAction(vaWriterMetadata,va)
          else if CompareOption(s,'wm2') then UpdateAction(vaWriterDetailedMetadata,va)
          else if CompareOption(s,'q') then UpdateAction(vaQuery,va)
          else if ReadOptionValue(s,'qx') then begin
            if UpdateAction(vaQuerySet,va) and TryStringToGUID(s,guid) then SnapID:=guid
            end
          else if ReadOptionValue(s,'qs') then begin
            if UpdateAction(vaQueryID,va) and TryStringToGUID(s,guid) then SnapID:=guid
            end
          else if CompareOption(s,'da') then UpdateAction(vaDeleteAll,va)
          else if ReadOptionValue(s,'dx') then begin
            if UpdateAction(vaDeleteSet,va) and TryStringToGUID(s,guid) then SnapID:=guid;
            end
          else if ReadOptionValue(s,'ds') then begin
            if UpdateAction(vaDeleteId,va) and TryStringToGUID(s,guid) then SnapID:=guid
            end
          else if ReadOptionValue(s,'log') then begin
            LogName:=s;
            WriteLn(Format('Writing log to file: "%s"',[s]));
            end
          end
        else begin
          s:=GetVolumeUniqueName(SetPathDelimiter(s));
          if length(s)>0 then VolList.Add(s);
          end;
        end;
      end;
    ok:=false;
    if va=vaHelp then ShowUsage
    else begin
      ok:=true;
      if length(ExecCommand)>0 then begin  // Check if the command is a valid CMD/EXE file
        dwAttributes:=FileGetAttr(ExecCommand);
        if (dwAttributes=faInvalid) or ((dwAttributes and faDirectory)<>0) then begin
          WriteLn(Format('ERROR: the parameter "%s" must be an existing file!',[execCommand]));
          WriteLn(' - Note: the -exec command cannot have parameters!');
          ok:=false;
          end;
        end;
      end;
    if ok then begin
      hr:=InitSecurity;
      if not succeeded(hr) then
        OleErrorHint(hr,Format(sLineBreak+'ERROR : COM call "%s" failed.',['CoInitializeSecurity']));
      VolumeShadowCopy:=TVolumeShadowCopy.Create;
      with VolumeShadowCopy do begin
        LogFilename:=LogName; //defLogFile;
        WriteLog:=true;
        if length(LogName)>0 then WriteLineToLog(WelcomeMsg);
        with IncludedWriterList do for i:=0 to Count-1 do
          WriteLineToLog(Format('(Option: Verifying inclusion of writer/component "%s")',[Strings[i]]));
        with ExcludedWriterList do for i:=0 to Count-1 do
          WriteLineToLog(Format('(Option: Excluding writer/component "%s")',[Strings[i]]));
        if Wait then WriteLineToLog('(Option: Wait on finish)');
        try
          case va of
          vaWriterStatus : begin   // List the writer state
                  WriteLineToLog('(Option: List writer status)');
                  Initialize(VSS_CTX_BACKUP);
                  // Gather writer metadata
                  GatherWriterMetadata;
                  // Gather writer status
                  GatherWriterStatus;
                  // List writer status
                  ListWriterStatus;
                  end;
          vaWriterMetadata : begin  // List the summary writer metadata
                  WriteLineToLog('(Option: List writer metadata)');
                  Initialize(VSS_CTX_BACKUP);
                  // Gather writer metadata
                  GatherWriterMetadata;
                  // List writer metadata
                  ListWriterMetadata(false);
                  end;
          vaWriterDetailedMetadata : begin  // List the extended writer metadata
                  WriteLineToLog('Option: List extended writer metadata');
                  Initialize(VSS_CTX_BACKUP);
                  // Gather writer metadata
                  GatherWriterMetadata;
                  // List writer metadata
                  ListWriterMetadata(true);
                  end;
          vaQuery : begin   // Query all shadow copies in the system
                  WriteLineToLog('(Option: Query all shadow copies)');
                  // Initialize the VSS client
                  Initialize(VSS_CTX_ALL);
                  // List all shadow copies in the system
                  QuerySnapshotSet(GUID_NULL);
                  end;
          vaQuerySet : begin  /// Query all shadow copies in the set
                  WriteLineToLog('(Option: Query shadow copy set');
                  // Initialize the VSS client
                  Initialize(VSS_CTX_ALL);
                  // List all shadow copies in the set
                  QuerySnapshotSet(SnapID);
                  end;
          vaQueryID : begin   // Query the specified shadow copy
                  WriteLineToLog('(Option: Query shadow copy)');
                  // Initialize the VSS client
                  Initialize(VSS_CTX_ALL);
                  // List all shadow copy properties
                  GetSnapshotProperties(SnapID);
                  end;
          vaDeleteAll : begin  // Delete all shadow copies in the system
                  WriteLineToLog('(Option: Delete all shadow copies)');
                  // Test if the user agrees
                  Write('This will delete all shadow copies in the system. Are you sure? [Y/N] ');
                  ReadLn(s);
                  if (Length(s)>0) and (UpCase(s[1])= 'Y') then begin
                    // Initialize the VSS client
                    Initialize(VSS_CTX_ALL);
                    // Delete all shadow copies in the system
                    DeleteAllSnapshots;
                    end;
                  end;
          vaDeleteSet : begin  // Delete a certain shadow copy set
                  WriteLineToLog('(Option: Delete a shadow copy set)');
                  // Initialize the VSS client
                  Initialize(VSS_CTX_ALL);
                  // Delete a certain shadow copy set
                  DeleteSnapshotSet(SnapID);
                  end;
          vaDeleteID : begin   // Delete a certain shadow copy
                  WriteLineToLog('(Option: Delete a shadow copy)');
                  // Initialize the VSS client
                  Initialize(VSS_CTX_ALL);
                  // Delete a certain shadow copy
                  DeleteSnapshot(SnapID);
                  end;
            else begin
              if VolList.Count>0 then begin
                if length(execCommand)> 0 then
                  WriteLineToLog(Format('(Option: Execute binary/script after shadow creation "%s")',[ExecCommand]));
                WriteLineToLog('(Option: Create shadow copy set)');
                // Initialize the VSS client
                Initialize(VSS_CTX_BACKUP);
                // Create the shadow copy set
                CreateSnapshotSet(VolList);
                // Execute BackupComplete, except in fast snapshot creation
                if VssContext and VSS_VOLSNAP_ATTR_DELAYED_POSTSNAPSHOT=0 then begin
                  try
                    // Executing the custom command if needed
                    if length(execCommand)> 0 then
                        ExecuteCommand(execCommand,GetLastShadowDeviceName);
                    ok:=true;
                  except
                    // Mark backup failure
                    ok:=false;
  //                  raise Exception.Create('Backup failed');
                    end;
                  // Complete the backup
                  // Note that this will notify writers that the backup is successful!
                  // (which means eventually log truncation)
                  if VssContext and VSS_VOLSNAP_ATTR_NO_WRITERS = 0 then BackupComplete(ok);
                  end;
                WriteLineToLog(sLineBreak+'Snapshot creation done.');
                end
              else WriteLn('Invalid volume');
              end;
            end;
          if Wait then begin
            Write('Strike <enter> to continue'); Readln;
            end;
        finally
          Free;
          end;
        end;
      end;
    VolList.Free; ExcludedWriterList.Free; IncludedWriterList.Free;
    end
  else ShowUsage;
  WriteLn('Done');
  CoUninitialize;
{$IFDEF DEBUG}
  readln;
{$EndIf}
  end.
