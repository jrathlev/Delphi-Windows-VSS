Notes on localization:
----------------------
Strings to be localized are located in this unit:
  VssConsts.pas:  (used in VssUtils.pas)

Included localizations:

Language    Subdirectory      Files   
-------------------------------------------
English     (default)         VssConst.pas
German      de                VssConst.pas

Sample snippet how to integrate VSS into own program:
-----------------------------------------------------
...
    procedure ShowVssStatus (const AStatus : string);
    begin
      ...
      end;
...
    VssThread:=CreateVssThread(Drive,true);
    with VssThread do begin
      LogFilename:=TempDir+'VssLog.txt';
      WriteLog:=true;
      OnStatusMessage:=ShowVssStatus;
      Resume;
      WriteLineToLog('Creating a Volume Shadow Copy: '+DateTimeToStr(Now));
      repeat
        Sleep(1);
        Application.ProcessMessages;
        until Done;
      if Success then begin
        SaveBackupComponentsDocument(TempDir+'VssBackupDoc.xml');
        SourceDrv:=ShadowDeviceName;
        end
      else begin
        WriteLineToLog(sLineBreak+'Snapshot creation failed');
        FreeAndNil(VssThread);
        SourceDrv:=Drive;
        end;
      end;
...
// Backup or similar action
...
  if assigned(VssThread) then begin
    try
      VssThread.DeleteShadowCopy;
    finally
      FreeAndNil(VssThread);
      end;
    end;
...

J. Rathlev, May 2017

