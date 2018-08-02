(* Delphi-Unit
   Routines and objects for Volume shadow copy service
   ===================================================

   Converted from the C++ program "vscsc" (volume shadow copy simple client)
   refer to: http://sourceforge.net/projects/vscsc/
   also refer to: WindowsSDK v7.0: Samples\winbase\vss\vshadow\

   © Dr. J. Rathlev, D-24222 Schwentinental (kontakt(a)rathlev-home.de)

   The contents of this file may be used under the terms of the
   Mozilla Public License ("MPL") or
   GNU Lesser General Public License Version 2 or later (the "LGPL")

   Software distributed under this License is distributed on an "AS IS" basis,
   WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for
   the specific language governing rights and limitations under the License.

   Version 1.0: November 2014
   Version 2.0: February 2015
                Thread based VSS, multilanguage support, bug fixes
           2.1: January 2016 - changes to use "resourcestring" for localization
           2.2: May 2017 - changes to log output (no localization)
           2.3: July 2017 - DiscoverDirectlyExcludedComponents fixed
   *)

unit VssUtils;

interface

uses
  System.SysUtils, System.Classes, System.Win.ComObj, System.Contnrs,
  Winapi.Windows, Winapi.ActiveX, WinApiUtils, VssApi;

const
  defLogFile : string = 'VssUtils.log';
  XmlSize = 512*1024;

type
// strings for status and log
// status strings: from resourcestring section for optional localization (VssConst.pas)
// log strings:    from const section - no localization
  TVssMessages = (nVssNotAvail,nInitBackupComps,nInitMetaData,nGatherMetaData,
    nGatherWriterstatus,nDscvDirExclComps,nDscvNsExclComps,nDscvAllExclComps,
    nDscvExclWriters,nDscvExpInclComps,nSelExpInclComps,nSaveBackCompsDoc,
    nAddVolSnapSet,nPrepBackup,nCreateShadowCopy,nCreateCopySuccess,
    nQueryShadowCopies,nQuerySnapshotSetID,nCreateShadowSet,nCompleteBackup);
  TLineBreak = (lbLeft,lbRight); // add line break to log string
  TLineBreaks = set of TLineBreak;

  TGuidList = array of TGuid;

  TOutputString = procedure(const s : string) of object;

  TGuidArray = array of TGuid;

  TLongWord = record
    case integer of
    0 : (LongWord : cardinal);
    1 : (Lo,Hi : word);
    2 : (LoL,LoH,HiL,HiH : byte);
    3 : (Bytes : array [0..3] of Byte);
    end;

  // The type of a file descriptor
  _VSS_DESCRIPTOR_TYPE = (VSS_FDT_UNDEFINED, VSS_FDT_EXCLUDE_FILES,
    VSS_FDT_FILELIST, VSS_FDT_DATABASE, VSS_FDT_DATABASE_LOG);

  TVssDescriptorType = _VSS_DESCRIPTOR_TYPE;

  ////////////////////////////////////////////////////////////////////////////////////
  TVssFileDescriptor = class(TObject)
  private
    path, filespec, alternatePath : string;
    isRecursive : boolean;
    fWriteLineToLog : TOutputString;
    function GetStringFromFileDescriptorType(AType : TVssDescriptorType) : string;
  public
    fdType : TVssDescriptorType;
    ExpandedPath, AffectedVolume : string;
    constructor Create;
    procedure Initialize(pFileDesc : IVssWMFiledesc;
      typeParam : TVssDescriptorType);
    procedure Print;
    property OnWriteLineToLog : TOutputString write fWriteLineToLog;
    end;

  /// /////////////////////////////////////////////////////////////////////////////////
  TVssComponent = class(TObject)
  private
    CompType : TVssComponentType;
    IsSelectable,NotifyOnBackupComplete,IsTopLevel,IsExcluded,IsExplicitlyIncluded : boolean;
    FullPath,CompName,LogicalPath,WriterName,Caption : string;
    Descriptors : TObjectList;
    AffectedPaths,AffectedVolumes : TStringList;
    fWriteLineToLog : TOutputString;
  public
    constructor Create;
    destructor Destroy; override;
    // Initialize from a IVssWMComponent
    procedure Initialize(const WriterNameParam : string; pComponent : IVssWMComponent); overload;
    // Initialize from a IVssComponent
    procedure Initialize(const WriterNameParam : string; pComponent : IVssComponent); overload;
    // Print summary/detalied information about this component
    procedure Print(bListDetailedInfo : boolean);
    // Convert a component type into a string
    function GetStringFromComponentType(eComponentType : TVssComponentType) : string;
    // Return TRUE if the current component is ancestor of the given component
    function IsAncestorOf(descendent : TVssComponent) : boolean;
    // return TRUEif it can be explicitly included
    function CanBeExplicitlyIncluded : boolean;
    property OnWriteLineToLog : TOutputString write fWriteLineToLog;
    end;

  /// /////////////////////////////////////////////////////////////////////////////////
  TVssWriter = class(TObject)
  private
    IsExcluded,supportsRestore,RebootRequiredAfterRestore : bool;
    RestoreMethod : VSS_RESTOREMETHOD_ENUM;
    WriterRestoreConditions : VSS_WRITERRESTORE_ENUM;
    ExcludedFiles,Components : TObjectList;
    InstanceId,WriterId,WriterName : string;
    fWriteLineToLog : TOutputString;
  public
    constructor Create;
    destructor Destroy; override;
    // Initialize from a IVssWMFiledesc
    procedure Initialize(Metadata : IVssExamineWriterMetadata);
    // Initialize from a IVssWriterComponentsExt
    procedure InitializeComponentsForRestore(WriterComponents : IVssWriterComponentsExt);
    // Print summary/detailed information about this writer
    procedure Print(bListDetailedInfo : boolean);
    function GetStringFromRestoreMethod(eRestoreMethod : TVssRestoreMethodEnum) : string;
    function GetStringFromRestoreConditions(eRestoreEnum : TVssWriterRestoreEnum) : string;
    property OnWriteLineToLog : TOutputString write fWriteLineToLog;
    end;

  /// /////////////////////////////////////////////////////////////////////////////////
  TVolumeShadowCopy = class(TObject)
  private
    CoInitializeCalled,     // TRUE if CoInitialize() was already called
                            // Needed to pair each successfull call to CoInitialize
                            // with a corresponding CoUninitialize
    DuringRestore,          // TRUE if we are during restore
    fWriteLog : boolean;    // write to log file
    LatestVolumeList : TStringList;
    // List of selected writers during the shadow copy creation process
    LatestSnapshotIdList : TGuidList;
    // List of shadow copy IDs from the latest shadow copy creation process
    LatestSnapshotSetID : TGuid; // Latest shadow copy set ID
    VssBackupComponents : IVssBackupComponents;
                        // The IVssBackupComponents interface
                        // is automatically released when this object is destructed.
                        // Needed to issue VSS calls
                        // List of resync pairs
    // map<VSS_ID,wstring,ltguid>      m_resyncPairs;
    fShowStatus : TOutputString;
    fLogFileName : string;
    LogFile : TextFile;
    WriterList : TObjectList;
    procedure WinCheck(Result : boolean; Hint : string = '');
    procedure WinCheckError(LastError : integer; Hint : string = '');
    procedure OleCheck(Result : HResult; Hint : string = '');
    procedure ShowStatusMsg(nMsg : TVssMessages; lbs : TLineBreaks = []);
    procedure ShowFormatStatusMsg(nMsg : TVssMessages; const Args: array of const; lbs : TLineBreaks = []);
    function DisplayNameForVolume(const volumeName : string) : string;
    procedure SetLogFilename(const Value : string);
    procedure SetWriteLog(Value : boolean);
    procedure PrintSnapshotProperties(Prop : TVssSnapshotProp);
    function GetStringFromWriterStatus(eWriterStatus : TVssWriterState) : string;
    procedure WaitAndCheckForAsyncOperation(Async : IVssAsync);
    function IsWriterSelected(guidInstanceId : TGuid) : boolean;
    procedure CheckSelectedWriterStatus;
    procedure DiscoverDirectlyExcludedComponents(ExcludedWriterAndComponentList : TStringList;
              AWriterList : TObjectList);
    procedure DiscoverNonShadowedExcludedComponents(ShadowSourceVolumes : TStringList);
    procedure DiscoverAllExcludedComponents;
    procedure DiscoverExcludedWriters;
    procedure DiscoverExplicitelyIncludedComponents;
    procedure VerifyExplicitelyIncludedWriter(const AWriterName : string;
              AWriterList : TObjectList);
    procedure VerifyExplicitelyIncludedComponent(const AIncludedComponent : string;
              AWriterList : TObjectList);
    procedure SelectExplicitelyIncludedComponents;
    procedure AddToSnapshotSet(VolumeList : TStringList);
    procedure PrepareForBackup;
    procedure DoSnapshotSet;
    procedure SetBackupSucceeded(succeeded : boolean);
  public
    VssContext : DWord; // VSS context
    constructor Create;
    destructor Destroy; override;
    procedure Initialize(AContext : DWord; AXmlDoc : string = '';
              ADuringRestore : boolean = false);
    procedure InitializeWriterMetadata;
    procedure GatherWriterMetadata;
    procedure GatherWriterStatus;
    procedure ListWriterMetadata(bListDetailedInfo : boolean);
    procedure ListWriterStatus;
    procedure SelectComponentsForBackup(ShadowSourceVolumes,ExcludedWriterAndComponentList,
              IncludedWriterAndComponentList : TStringList);
    procedure CreateSnapshotSet(AVolumeList : TStringList; AExcludedWriterList : TStringList = nil;
              AIncludedWriterList : TStringList = nil; AOutXmlFile : string = ''); overload;
    procedure CreateSnapshotSet(const AVolume : string); overload;
    procedure BackupComplete(succeeded : boolean);
    function GetLastShadowDeviceName : string;
    procedure ExecuteCommand(const Command,Param : string);
    procedure QuerySnapshotSet(SnapshotSetID : TGuid);
    procedure GetSnapshotProperties(SnapshotSetID : TGuid);
    procedure DeleteAllSnapshots;
    procedure DeleteSnapshotSet(snapshotSetID : TGuid);
    procedure DeleteOldestSnapshot(const stringVolumeName : string);
    procedure DeleteSnapshot(snapshotID : TGuid);
    procedure SaveBackupComponentsDocument(const AFileName : string);

    procedure WriteLineToLog(const s : string);
    property LogFileName : string read fLogFileName write SetLogFilename;
    property WriteLog : boolean read fWriteLog write SetWriteLog;
    property OnStatusMessage : TOutputString write fShowStatus;
    end;

  { ---------------------------------------------------------------- }
  TVssThread = class (TThread)
  private
    FVolumeShadowCopy : TVolumeShadowCopy;
    FDrive,
    FShadowDeviceName : string;
    FSuccess : boolean;
    function GetLogFileName : string;
    procedure SetLogFileName (const Value : string);
    function GetWriteLog : boolean;
    procedure SetWriteLog (Value : boolean);
    procedure SetShowStatus (ACallBack : TOutputString);
    function GetDone : boolean;
  protected
    procedure Execute; override;
  public
    constructor Create (const ADrive : string; ASuspend : Boolean = true;
                        APriority : TThreadPriority = tpNormal);
    destructor Destroy; override;
    procedure WriteLineToLog (const  s : string);
    procedure SaveBackupComponentsDocument(const AFileName : string);
    procedure DeleteShadowCopy;
    property LogFileName : string read GetLogFileName write SetLogFilename;
    property WriteLog : boolean read GetWriteLog write SetWriteLog;
    property Done  : boolean read GetDone;
    property Success : boolean read FSuccess;
    property ShadowDeviceName : string read FShadowDeviceName;
    property OnStatusMessage : TOutputString write SetShowStatus;
    end;

  EVssError = class(Exception);

  { ---------------------------------------------------------------- }
// Convert VSS time to DateTime string
function VssTimeToString(vssTime : int64) : string;

// add "\" only if APath is not empty
function SetPathDelimiter(const APath : string) : string;

// Get the unique volume name for the given path without throwing an error
function GetUniqueVolumeNameForPath(path : string; var volname : string) : boolean;

// Get the displayable root path for the given volume name
function GetDisplayNameForVolume(const volumeName : string; var volumeNameCanon : string) : boolean;

// Convert a failure type into a string
function HResultToString(hrError : HResult) : string;

// Raise EOleSysError exception from an error code and show hint
procedure OleErrorHint(ErrorCode : HResult; const Hint : string = '');

// Add Guid to list
procedure AddGuidToList(var AGuidList : TGuidList; AGuid : TGuid);

// Convert a string to a GUID, return false if failed
function TryStringToGUID(const S: string; var AGuid: TGUID) : boolean;

// Utility function to write a new file
procedure WriteStringToFile(const AFileName,bstrXML : string);

{ ---------------------------------------------------------------- }
// Initialize COM security
function InitSecurity : HResult;

// Create VSS thread
function CreateVssThread (const ADrive : string; IgnoreTooLate : boolean = false) : TVssThread;

{ ---------------------------------------------------------------- }
implementation

uses System.StrUtils, VssConsts;

const
  cpUtf8 = 65001;

  // do not localize
  sVssNotAvail = 'Volume Shadow Copy could not be initialized on this system!';
  sInitBackupComps = 'Initializing IVssBackupComponents Interface ...';
  sInitMetaData = 'Initialize writer metadata ...';
  sGatherMetaData = 'Gathering writer metadata ...';
  sGatherWriterstatus = 'Gathering writer status ...';
  sDscvDirExclComps = 'Discover directly excluded components ...';
  sDscvNsExclComps = 'Discover components that reside outside the shadow set ...';
  sDscvAllExclComps = 'Discover all excluded components ...';
  sDscvExclWriters = 'Discover excluded writes ...';
  sDscvExpInclComps = 'Discover explicitly included components ...';
  sSelExpInclComps = 'Select explicitly included components ...';
  sSaveBackCompsDoc = 'Saving the backup components document ...';
  sAddVolSnapSet = 'Add volumes to snapshot set ...';
  sPrepBackup = 'Preparing for backup ...';
  sCreateShadowCopy = 'Creating the shadow copy ...';
  sCreateCopySuccess = 'Shadow copy set successfully created';
  sQueryShadowCopies = 'Querying all shadow copies in the system ...';
  sQuerySnapshotSetID = 'Querying all shadow copies with the SnapshotSetID %s ...';
  sCreateShadowSet = 'Creating shadow set {%s} ...';
  sCompleteBackup = 'Completing the backup ...';

  VssMessageStrings : array [TVssMessages] of string = (
    sVssNotAvail       ,
    sInitBackupComps   ,
    sInitMetaData      ,
    sGatherMetaData    ,
    sGatherWriterStatus,
    sDscvDirExclComps  ,
    sDscvNsExclComps   ,
    sDscvAllExclComps  ,
    sDscvExclWriters   ,
    sDscvExpInclComps  ,
    sSelExpInclComps   ,
    sSaveBackCompsDoc  ,
    sAddVolSnapSet     ,
    sPrepBackup        ,
    sCreateShadowCopy  ,
    sCreateCopySuccess ,
    sQueryShadowCopies ,
    sQuerySnapshotSetID,
    sCreateShadowSet   ,
    sCompleteBackup    );

  VssResourceStrings : array [TVssMessages] of PResStringRec = (
    @rsVssNotAvail       ,
    @rsInitBackupComps   ,
    @rsInitMetaData      ,
    @rsGatherMetaData    ,
    @rsGatherWriterStatus,
    @rsDscvDirExclComps  ,
    @rsDscvNsExclComps   ,
    @rsDscvAllExclComps  ,
    @rsDscvExclWriters   ,
    @rsDscvExpInclComps  ,
    @rsSelExpInclComps   ,
    @rsSaveBackCompsDoc  ,
    @rsAddVolSnapSet     ,
    @rsPrepBackup        ,
    @rsCreateShadowCopy  ,
    @rsCreateCopySuccess ,
    @rsQueryShadowCopies ,
    @rsQuerySnapshotSetID,
    @rsCreateShadowSet   ,
    @rsCompleteBackup    );

{ ---------------------------------------------------------------- }
// Convert boolean to adjustable string
function BooleanToString(B : boolean; const sTrue : string = 'TRUE';
  sFalse : string = 'FALSE') : string;
begin
  if B then Result:=sTrue
  else Result:=sFalse;
  end;

// Convert VSS time to DateTime string
function VssTimeToString(vssTime : int64) : string;
var
  stLocal : TSystemTime;
  ftLocal : TFileTime;
begin
  // Compensate for local TZ
  FileTimeToLocalFileTime(TFileTime(vssTime),ftLocal);
  // Finally convert it to system time
  FileTimeToSystemTime(ftLocal,stLocal);
  Result:=FormatDateTime('',SystemTimeToDateTime(stLocal));
  end;

// add "\" only if APath is not empty
function SetPathDelimiter(const APath : string) : string;
begin
  if length(APath)>0 then Result:=IncludeTrailingPathDelimiter(APath)
  else Result:='';
  end;


function BStrToString(bstr : PWideChar) : string;
begin
  if bstr=nil then Result:='' else Result:=bstr;
  end;

// Get the unique volume name for the given path without throwing an error
function GetUniqueVolumeNameForPath(path : string; var volname : string) : boolean;
var
  uname : array [0 .. MAX_PATH] of WideChar;
begin
  Result:=false;
  if length(path)=0 then Exit;
  // Add the backslash termination, if needed
  path:=SetPathDelimiter(path);
  if ClusterIsPathOnSharedVolume(PWideChar(path)) then begin
    Result:=ClusterGetVolumePathName(PWideChar(path),uname,MAX_PATH+1);
    if Result then volname:=uname
    else volname:='';
    end
  else begin
    // Get the root path of the volume
    Result:=GetVolumePathName(PWideChar(path),uname,MAX_PATH+1);
    if Result then
    begin
      volname:=GetVolumeUniqueName(uname);
      Result:=length(volname) > 0;
      end;
    end;
  end;

// Get the displayable root path for the given volume name
function GetDisplayNameForVolume(const volumeName : string;
  var volumeNameCanon : string) : boolean;
var
  dwRequired : cardinal;
  volumeMountPoints : array of WideChar;
  i,n : integer;
begin
  Result:=false;
  SetLength(volumeMountPoints,MAX_PATH);
  if not GetVolumePathNamesForVolumeName(PWideChar(volumeName),
      @volumeMountPoints[0],length(volumeMountPoints),dwRequired) then begin
    // If not enough, retry with a larger size
    SetLength(volumeMountPoints,dwRequired);
    if (dwRequired = 0) or not GetVolumePathNamesForVolumeName
      (PWideChar(volumeName),@volumeMountPoints[0],length(volumeMountPoints),
      dwRequired) then begin
        volumeMountPoints:=nil;
        Exit;
        end;
    end;
  // compute the smallest mount point by enumerating the returned MULTI_SZ
  SetLength(volumeNameCanon,dwRequired);
  i:=0;
  repeat
    if volumeMountPoints[i] <> #0 then begin
      n:=i;
      repeat
        inc(i)
      until volumeMountPoints[i] = #0;
      if (i-n) < length(volumeNameCanon) then begin
        Move(volumeMountPoints[n],volumeNameCanon[1],
          (i-n+1) * sizeof(WideChar));
        SetLength(volumeNameCanon,i-n);
        end;
      inc(i);
      end
    else inc(i);
  until volumeMountPoints[i] = #0;
  volumeMountPoints:=nil;
  Result:=true;
  end;

// Convert a failure type into a string
function HResultToString(hrError : HResult) : string;
begin
  case hrError of
    VSS_E_BAD_STATE:             Result:='VSS: called function was in incorrect state';
    VSS_E_UNEXPECTED:            Result:='VSS: an unexpected error was encountered';
    VSS_E_PROVIDER_ALREADY_REGISTERED: Result:='VSS: provider has already been registered';
    VSS_E_PROVIDER_NOT_REGISTERED:     Result:='VSS: provider is not registered';
    VSS_E_PROVIDER_VETO:         Result:='VSS: provider had an error';
    VSS_E_PROVIDER_IN_USE:       Result:='VSS: provider is currently in use';
    VSS_E_OBJECT_NOT_FOUND:      Result:='VSS: specified object was not found';
    VSS_S_ASYNC_PENDING:         Result:='VSS: asynchronous operation is pending';
    VSS_S_ASYNC_FINISHED:        Result:='VSS: asynchronous operation has completed';
    VSS_S_ASYNC_CANCELLED:       Result:='VSS: asynchronous operation has been cancelled';
    VSS_E_VOLUME_NOT_SUPPORTED:  Result:='VSS: specified volume is not supported';
    VSS_E_VOLUME_NOT_SUPPORTED_BY_PROVIDER: Result:='VSS: provider does not support the specified volume';
    VSS_E_OBJECT_ALREADY_EXISTS: Result:='VSS: object already exists.';
    VSS_E_UNEXPECTED_PROVIDER_ERROR: Result:='VSS:  unexpected provider error';
    VSS_E_CORRUPT_XML_DOCUMENT,
    VSS_E_INVALID_XML_DOCUMENT:  Result:='VSS: XML document is invalid';
    VSS_E_MAXIMUM_NUMBER_OF_VOLUMES_REACHED: Result:='VSS: maximum number of volumes reached';
    VSS_E_FLUSH_WRITES_TIMEOUT:  Result:='VSS: provider timed out while flushing data';
    VSS_E_HOLD_WRITES_TIMEOUT:   Result:='VSS: provider timed out while holding writes';
    VSS_E_UNEXPECTED_WRITER_ERROR: Result:='VSS: problems while sending events to writers';
    VSS_E_SNAPSHOT_SET_IN_PROGRESS: Result:='VSS: Another shadow copy creation is already in progress';
    VSS_E_MAXIMUM_NUMBER_OF_SNAPSHOTS_REACHED: Result:='VSS: maximum number of shadow copies reached';
    VSS_E_WRITER_INFRASTRUCTURE: Result:='VSS: problem occurred while trying to contact writers';
    VSS_E_WRITER_NOT_RESPONDING: Result:='VSS: writer did not respond to a GatherWriterStatus';
    VSS_E_WRITER_ALREADY_SUBSCRIBED: Result:='VSS: writer has already successfully called the Subscribe function';
    VSS_E_UNSUPPORTED_CONTEXT:   Result:='VSS: provider does not support the specified shadow copy type';
    VSS_E_VOLUME_IN_USE:         Result:='VSS: specified shadow copy storage association is in use';
    VSS_E_MAXIMUM_DIFFAREA_ASSOCIATIONS_REACHED: Result:='VSS: Maximum number of shadow copy storage associations already reached';
    VSS_E_INSUFFICIENT_STORAGE:  Result:='VSS: Insufficient storage available';
    VSS_E_NO_SNAPSHOTS_IMPORTED: Result:='VSS: No shadow copies were successfully imported';
    VSS_S_SOME_SNAPSHOTS_NOT_IMPORTED: Result:='VSS: Some shadow copies were not successfully imported';
    VSS_E_MAXIMUM_NUMBER_OF_REMOTE_MACHINES_REACHED: Result:='VSS: maximum number of remote machines has been reached.';
    VSS_E_REMOTE_SERVER_UNAVAILABLE: Result:='VSS: remote server is unavailable';
    VSS_E_REMOTE_SERVER_UNSUPPORTED: Result:='VSS: that does not support remote shadow-copy creation';
    VSS_E_REVERT_IN_PROGRESS:    Result:='VSS: revert is currently in progress';
    VSS_E_REVERT_VOLUME_LOST:    Result:='VSS: volume being reverted was lost during revert';
    VSS_E_REBOOT_REQUIRED:       Result:='VSS: reboot is required after this operation';
    VSS_E_TRANSACTION_FREEZE_TIMEOUT: Result:='VSS: timeout occurred while freezing a transaction manager';
    VSS_E_TRANSACTION_THAW_TIMEOUT:  Result:='VSS: Too much time elapsed between freezing and thawing';
    VSS_E_VOLUME_NOT_LOCAL:      Result:='VSS: volume is not mounted on the local host';
    VSS_E_CLUSTER_TIMEOUT:       Result:='VSS: timeout occurred while preparing a cluster shared volume';
    VSS_E_WRITERERROR_INCONSISTENTSNAPSHOT: Result:='VSS: shadow-copy set contains only a subset of the volumes';
    VSS_E_WRITERERROR_OUTOFRESOURCES: Result:='VSS: resource allocation failed';
    VSS_E_WRITERERROR_TIMEOUT:   Result:='VSS: timeout expired between Freeze and Thaw';
    VSS_E_WRITERERROR_RETRYABLE: Result:='VSS: writer experienced a transient error';
    VSS_E_WRITERERROR_NONRETRYABLE: Result:='VSS: writer experienced a non-transient error';
    VSS_E_WRITERERROR_RECOVERY_FAILED: Result:='VSS: writer experienced an error while trying to recover';
    VSS_E_BREAK_REVERT_ID_FAILED: Result:='VSS: break operation failed, partition identities could not be reverted';
    VSS_E_LEGACY_PROVIDER:       Result:='VSS: version of hardware provider does not support this operation';
    VSS_E_MISSING_DISK:          Result:='VSS: expected disk did not arrive in the system';
    VSS_E_MISSING_HIDDEN_VOLUME: Result:='VSS: expected hidden volume did not arrive in the system';
    VSS_E_MISSING_VOLUME:        Result:='VSS: expected volume did not arrive in the system';
    VSS_E_AUTORECOVERY_FAILED:   Result:='VSS: autorecovery operation failed';
    VSS_E_DYNAMIC_DISK_ERROR:    Result:='VSS: error processing the dynamic disks';
    VSS_E_NONTRANSPORTABLE_BCD:  Result:='VSS: Backup Components Document is non-transportable';
    VSS_E_CANNOT_REVERT_DISKID:  Result:='VSS: MBR signature or GPT ID could not be set';
    VSS_E_RESYNC_IN_PROGRESS:    Result:='VSS: LUN resynchronization operation could not be started';
    VSS_E_CLUSTER_ERROR:         Result:='VSS: clustered disks could not be enumerated';
    VSS_E_ASRERROR_DISK_ASSIGNMENT_FAILED: Result:='VSS: too few disks on this computer';
    VSS_E_ASRERROR_DISK_RECREATION_FAILED: Result:='VSS: cannot create a disk on this computer needed to restore';
    VSS_E_ASRERROR_NO_ARCPATH:   Result:='VSS: computer needs to be restarted to finish preparing a hard disk for restore';
    VSS_E_ASRERROR_MISSING_DYNDISK: Result:='VSS: backup failed due to a missing disk for a dynamic volume';
    VSS_E_ASRERROR_SHARED_CRIDISK: Result:='VSS: Automated System Recovery failed';
    VSS_E_ASRERROR_DATADISK_RDISK0: Result:='VSS: data disk is currently set as active in BIOS.';
    VSS_E_ASRERROR_RDISK0_TOOSMALL: Result:='VSS: disk that is set as active in BIOS is too small';
    VSS_E_ASRERROR_CRITICAL_DISKS_TOO_SMALL: Result:='VSS: failed to find enough suitable disks';
    VSS_E_WRITER_STATUS_NOT_AVAILABLE: Result:='VSS: writer status is not available';
    VSS_E_UNSELECTED_VOLUME:     Result:='VSS: operation would overwrite a volume that is not explicitly selected';
    VSS_E_SNAPSHOT_NOT_IN_SET:   Result:='VSS: shadow copy ID was not found';
    VSS_E_NESTED_VOLUME_LIMIT:   Result:='VSS: specified volume is nested too deeply';
    VSS_E_ASRERROR_DYNAMIC_VHD_NOT_SUPPORTED: Result:='critical dynamic disk is a Virtual Hard Disk';
    VSS_E_CRITICAL_VOLUME_ON_INVALID_DISK:    Result:='critical volume exists on a disk which cannot be backed up by ASR';
    VSS_E_ASRERROR_RDISK_FOR_SYSTEM_DISK_NOT_FOUND: Result:='No disk that can be used for recovering the system disk found';
    VSS_E_ASRERROR_NO_PHYSICAL_DISK_AVAILABLE: Result:='no fixed disk found that can be used to recreate volumes';
    VSS_E_ASRERROR_FIXED_PHYSICAL_DISK_AVAILABLE_AFTER_DISK_EXCLUSION: Result:='no disk found which it can use for recreating volumes';
    VSS_E_ASRERROR_CRITICAL_DISK_CANNOT_BE_EXCLUDED: Result:='Restore failed because a disk which was critical at backup is excluded';
    VSS_E_ASRERROR_SYSTEM_PARTITION_HIDDEN: Result:='active system partition is hidden';
    // Regular COM errors
    S_OK:                        Result:='Operation succeeded';
    S_FALSE:                     Result:='Operation succeeded but returned "FALSE"';
    E_UNEXPECTED:                Result:='Catastrophic failure';
    E_OUTOFMEMORY:               Result:='Ran out of memory';
    E_NOTIMPL:                   Result:='Not implemented';
    RPC_E_TOO_LATE:              Result:='Security must be initialized before any interfaces are marshalled or unmarshalled. It cannot be changed once initialized.';
    RPC_E_NO_GOOD_SECURITY_PACKAGES: Result:='No security packages are installed on this machine or the user is not logged on or there are no compatible security packages between the client and server.';
    else begin
      if TLongWord(hrError).Hi and $7FF = FACILITY_WIN32 then begin
        Result:=SysErrorMessage(hrError and $FFFF);
        end
      else Result:='<Unknown error code>';
      end;
    end;
  end;

// Raise EOleSysError exception from an error code and show hint
procedure OleErrorHint(ErrorCode : HResult; const Hint : string = '');
begin
  raise EOleSysError.Create(Hint+' ('+HResultToString(ErrorCode)+')',ErrorCode,0);
  end;

// Raise EOleSysError exception if result code indicates an error }
procedure OleCheck(Result : HResult; AWriteLineToLog : TOutputString; Hint : string = '');
begin
  if not succeeded(Result) then begin
    if length(Hint) = 0 then Hint:='<unknown>';
    if assigned(AWriteLineToLog) then begin
      AWriteLineToLog(Format('ERROR : COM call "%s" failed.',[Hint]));
      AWriteLineToLog(Format('- Returned HRESULT = $%.8x',[Result]));
      AWriteLineToLog(Format('- Error text: %s',[HResultToString(Result)]));
      end;
    OleErrorHint(Result,Hint);
    end;
  end;

// Add Guid to list
procedure AddGuidToList(var AGuidList : TGuidList; AGuid : TGuid);
begin
  SetLength(AGuidList,length(AGuidList)+1);
  AGuidList[High(AGuidList)]:=AGuid;
  end;

// Convert a string to a GUID, return false if failed
function TryStringToGUID(const S: string; var AGuid: TGUID) : boolean;
begin
  Result:=Succeeded(CLSIDFromString(PWideChar(S),AGuid));
  end;

// Utility function to write a new file
procedure WriteStringToFile(const AFileName,bstrXML : string);
var
  fs : TFileStream;
begin
  fs:=TFileStream.Create(AFileName,fmCreate);
  with fs do begin
    Write(bstrXML[1],2*length(bstrXML));
    Free;
    end;
  end;

// Open log file (create or append)
function AppendFile(FName : string; var LogFile : TextFile) : boolean;
begin
  if FileExists(FName) then begin
    assignfile(LogFile,FName,cpUtf8);
    {$I-} System.append(LogFile) {$I+};
    Result:=IOResult=0;
    if Result then WriteLn(LogFile);
    end
  else begin
    assignfile(LogFile,FName,cpUtf8);
    {$I-} System.rewrite(LogFile) {$I+};
    Result:=IOResult=0;
    end;
  end;

/// /////////////////////////////////////////////////////////////////////////////////
// VssFileDescriptor
//
constructor TVssFileDescriptor.Create;
begin
  inherited Create;
  isRecursive:=false;
  fWriteLineToLog:=nil;
  end;

procedure TVssFileDescriptor.Initialize(pFileDesc : IVssWMFiledesc;
  typeParam : TVssDescriptorType);
var
  bstrPath,bstrFilespec,bstrAlternate : PWideChar;
  bRecursive : bool;
  n : integer;
begin
  // Set the type
  fdType:=typeParam;
  OleCheck(pFileDesc.GetPath(bstrPath),fWriteLineToLog,
    'TVssFileDescriptor.Initialize:pFileDesc.GetPath');
  OleCheck(pFileDesc.GetFilespec(bstrFilespec),fWriteLineToLog,
    'TVssFileDescriptor.Initialize:pFileDesc.GetFilespec');
  OleCheck(pFileDesc.GetRecursive(bRecursive),fWriteLineToLog,
    'TVssFileDescriptor.Initialize:pFileDesc.GetRecursive');
  OleCheck(pFileDesc.GetAlternateLocation(bstrAlternate),fWriteLineToLog,
    'TVssFileDescriptor.Initialize:pFileDesc.GetAlternateLocation');
  // Initialize local data members
  path:=OleStrToString(bstrPath);
  SysFreeString(bstrPath);
  filespec:=OleStrToString(bstrFilespec);
  SysFreeString(bstrFilespec);
  alternatePath:=OleStrToString(bstrAlternate);
  SysFreeString(bstrAlternate);
  isRecursive:=bRecursive;
  if length(path)=0 then Exit;
  // Get the expanded path
  SetLength(ExpandedPath,MAX_PATH);
  n:=ExpandEnvironmentStrings(PWideChar(path),PWideChar(ExpandedPath),MAX_PATH);
  if n>MAX_PATH then begin
    SetLength(ExpandedPath,n);
    n:=ExpandEnvironmentStrings(PWideChar(path),PWideChar(ExpandedPath),MAX_PATH);
    end;
  SetLength(ExpandedPath,n-1);
  ExpandedPath:=SetPathDelimiter(ExpandedPath);
  // Get the affected volume
  if (not GetUniqueVolumeNameForPath(ExpandedPath,AffectedVolume)) then
    AffectedVolume:=ExpandedPath;
  end;

function TVssFileDescriptor.GetStringFromFileDescriptorType
  (AType : TVssDescriptorType) : string;
begin
  case AType of
    VSS_FDT_UNDEFINED:        Result:='Undefined';
    VSS_FDT_EXCLUDE_FILES:    Result:='Exclude';
    VSS_FDT_FILELIST:         Result:='File List';
    VSS_FDT_DATABASE:         Result:='Database';
    VSS_FDT_DATABASE_LOG:     Result:='Database Log';
    else Result:='Undefined';
    end;
  end;

procedure TVssFileDescriptor.Print;
var
  s : string;
begin
  // Print writer identity information
  if assigned(fWriteLineToLog) then begin
    if length(alternatePath) > 0 then s:=' ,Alternate Location = '+alternatePath
    else s:='';
    if isRecursive then s:=' Recursive : '+BooleanToString(isRecursive,'Yes','No')+s;
    fWriteLineToLog(Format('      - %s : Path = %s, Filespec = %s%s',
      [GetStringFromFileDescriptorType(fdType),path,filespec,s]));
    end;
  end;

/// /////////////////////////////////////////////////////////////////////////////////
// VssComponent
//
// In-memory representation of a component
constructor TVssComponent.Create;
begin
  inherited Create;
  CompType:=VSS_CT_UNDEFINED;
  IsSelectable:=false;
  NotifyOnBackupComplete:=false;
  IsTopLevel:=false;
  IsExcluded:=false;
  IsExplicitlyIncluded:=false;
  Descriptors:=TObjectList.Create;
  AffectedPaths:=TStringList.Create;
  AffectedPaths.Sorted:=true;
  AffectedVolumes:=TStringList.Create;
  AffectedVolumes.Sorted:=true;
  OnWriteLineToLog:=nil;
  end;

destructor TVssComponent.Destroy;
begin
  Descriptors.Free;
  AffectedPaths.Free;
  AffectedVolumes.Free;
  inherited Destroy;
  end;

// Initialize from a IVssWMComponent
procedure TVssComponent.Initialize(const WriterNameParam : string; pComponent : IVssWMComponent);
var
  pInfo : PVssComponentInfo;
  i : integer;
  pFileDesc : IVssWMFiledesc;
  desc : TVssFileDescriptor;
begin
  WriterName:=WriterNameParam;
  // Get the component info
  OleCheck(pComponent.GetComponentInfo(pInfo),fWriteLineToLog,
    'TVssComponent.Initialize:pComponent.GetComponentInfo');
  // Initialize local members
  with pInfo^ do begin
    CompName:=BstrToString(bstrComponentName);
    LogicalPath:=BstrToString(bstrLogicalPath);
    Caption:=BstrToString(bstrCaption);
    CompType:=TVssComponentType(ciType);
    IsSelectable:=bSelectable;
    NotifyOnBackupComplete:=bNotifyOnBackupComplete;
    end;
  // Compute the full path
  FullPath:=SetPathDelimiter(LogicalPath)+CompName;
  if not AnsiStartsText('\',FullPath) then FullPath:='\'+FullPath;
  with pInfo^ do begin
    // Get file list descriptors
    if cFileCount>0 then for i:=0 to cFileCount-1 do begin
      OleCheck(pComponent.GetFile(i,pFileDesc),fWriteLineToLog,
        'TVssComponent.Initialize:pComponent.GetFile');
      desc:=TVssFileDescriptor.Create;
      desc.OnWriteLineToLog:=fWriteLineToLog;
      desc.Initialize(pFileDesc,VSS_FDT_FILELIST);
      Descriptors.Add(desc);
      end;
    // Get database descriptors
    if cDatabases>0 then for i:=0 to cDatabases-1 do begin
      OleCheck(pComponent.GetDatabaseFile(i,pFileDesc),fWriteLineToLog,
        'TVssComponent.Initialize:pComponent.GetDatabaseFile');
      desc:=TVssFileDescriptor.Create;
      desc.OnWriteLineToLog:=fWriteLineToLog;
      desc.Initialize(pFileDesc,VSS_FDT_DATABASE);
      Descriptors.Add(desc);
      end;
    // Get log descriptors
    if cLogFiles>0 then for i:=0 to cLogFiles-1 do begin
      OleCheck(pComponent.GetDatabaseLogFile(i,pFileDesc),fWriteLineToLog,
        'TVssComponent.Initialize:pComponent.GetDatabaseLogFile');
      desc:=TVssFileDescriptor.Create;
      desc.OnWriteLineToLog:=fWriteLineToLog;
      desc.Initialize(pFileDesc,VSS_FDT_DATABASE_LOG);
      Descriptors.Add(desc);
      end;
    end;
  pComponent.FreeComponentInfo(pInfo);
  // Compute the affected paths and volumes
  for i:=0 to Descriptors.Count-1 do with Descriptors[i] as TVssFileDescriptor do begin
    if AffectedPaths.IndexOf(ExpandedPath) < 0 then AffectedPaths.Add(ExpandedPath);
    if AffectedVolumes.IndexOf(AffectedVolume) < 0 then AffectedVolumes.Add(AffectedVolume);
    end;
//  Print(true);
  end;

// Initialize from a IVssComponent
procedure TVssComponent.Initialize(const WriterNameParam : string; pComponent : IVssComponent);
var
  bstrComponentName,bstrLogicalPath : PWideChar;
begin
  WriterName:=WriterNameParam;
  // Get component type
  OleCheck(pComponent.GetComponentType(CompType),fWriteLineToLog,
    'TVssComponent.Initialize:pComponent.GetComponentType');
  // Get component name
  OleCheck(pComponent.GetComponentName(bstrComponentName),fWriteLineToLog,
    'TVssComponent.Initialize:pComponent.GetComponentName');
  CompName:=OleStrToString(bstrComponentName);
  SysFreeString(bstrComponentName);
  // Get component logical path
  OleCheck(pComponent.GetLogicalPath(bstrLogicalPath),fWriteLineToLog,
    'TVssComponent.Initialize:pComponent.GetLogicalPath');
  LogicalPath:=OleStrToString(bstrLogicalPath);
  SysFreeString(bstrLogicalPath);
  // Compute the full path
  FullPath:=SetPathDelimiter(LogicalPath)+CompName;
  if not AnsiStartsText('\',FullPath) then FullPath:='\'+FullPath;
  end;

// Print summary/detailed information about this component
procedure TVssComponent.Print(bListDetailedInfo : boolean);
var
  i : integer;
  wsLocalVolume : string;
begin
  // Print writer identity information
  if assigned(fWriteLineToLog) then begin
    fWriteLineToLog(Format('    - Component: "%s:%s"'+sLineBreak+
            '      - Name: "%s"'+sLineBreak+
            '      - Logical Path: "%s"'+sLineBreak+
            '      - Full Path: "%s"'+sLineBreak+
            '      - Caption: "%s"'+sLineBreak+
            '      - Type: %s [%d]'+sLineBreak+
            '      - Is Selectable: %s'+sLineBreak +
            '      - Is top level: %s'+sLineBreak +
            '      - Notify on backup complete: %s',
      [WriterName,FullPath,CompName,LogicalPath,FullPath,Caption,
       GetStringFromComponentType(CompType),integer(CompType),
       BooleanToString(IsSelectable,'Yes','No'),BooleanToString(IsTopLevel,'Yes','No'),
       BooleanToString(NotifyOnBackupComplete,'Yes','No')]));
    // Compute the affected paths and volumes
    if bListDetailedInfo then begin
      fWriteLineToLog(Format('      - Components (%u):',[Descriptors.Count]));
      for i:=0 to Descriptors.Count-1 do (Descriptors[i] as TVssFileDescriptor).Print;
      end;
    // Print the affected paths and volumes
    fWriteLineToLog(Format('        - Affected paths by this component (%u):',[AffectedPaths.Count]));
    for i:=0 to AffectedPaths.Count-1 do fWriteLineToLog(Format('          - %s',[AffectedPaths[i]]));
    fWriteLineToLog(Format('        - Affected volumes by this component (%u):',[AffectedVolumes.Count]));
    for i:=0 to AffectedVolumes.Count-1 do begin
      if GetDisplayNameForVolume(AffectedVolumes[i],wsLocalVolume) then
           fWriteLineToLog(Format('          - %s [%s]',[AffectedVolumes[i],wsLocalVolume]))
      else fWriteLineToLog(Format('          - %s [Not valid for local machine]',[AffectedVolumes[i]]));
      end;
    end;
  end;

// Convert a component type into a string
function TVssComponent.GetStringFromComponentType(eComponentType : TVssComponentType) : string;
begin
  case eComponentType of
    VSS_CT_DATABASE:    Result:='VSS_CT_DATABASE';
    VSS_CT_FILEGROUP:   Result:='VSS_CT_FILEGROUP';
    else Result:='Unknown constant';
    end;
  end;

// Return TRUE if the current component is ancestor of the given component
function TVssComponent.IsAncestorOf(descendent : TVssComponent) : boolean;
var
  fullPathAppendedWithBackslash,descendentPathAppendedWithBackslash : string;
begin
  // The child must have a longer full path
  if length(descendent.FullPath) <= length(FullPath) then Result:=false
  else begin
    fullPathAppendedWithBackslash:=SetPathDelimiter(FullPath);
    descendentPathAppendedWithBackslash:=SetPathDelimiter(descendent.FullPath);
    // Return TRUE if the current full path is a prefix of the child full path
    Result:=AnsiStartsText(fullPathAppendedWithBackslash,descendentPathAppendedWithBackslash);
    end;
  end;

// return TRUEif it can be explicitly included
function TVssComponent.CanBeExplicitlyIncluded : boolean;
begin
  if IsExcluded then Result:=false
    // selectable can be explictly included
  else if IsSelectable then Result:=true
    // Non-selectable top level can be explictly included
  else if IsTopLevel then Result:=true
  else Result:=false;
  end;

/// /////////////////////////////////////////////////////////////////////////////////
// VssWriter
//
constructor TVssWriter.Create;
begin
  inherited Create;
  IsExcluded:=false;
  supportsRestore:=false;
  RebootRequiredAfterRestore:=false;
  RestoreMethod:=VSS_RME_UNDEFINED;
  WriterRestoreConditions:=VSS_WRE_UNDEFINED;
  ExcludedFiles:=TObjectList.Create;
  Components:=TObjectList.Create;
  fWriteLineToLog:=nil;
  end;

destructor TVssWriter.Destroy;
begin
  ExcludedFiles.Free;
  Components.Free;
  inherited Destroy;
  end;

// Initialize from a IVssWMFiledesc
procedure TVssWriter.Initialize(Metadata : IVssExamineWriterMetadata);
var
  idInstance,idWriter : TGuid;
  bstrWriterName,bstrService,bstrUserProcedure : PWideChar;
  Usage : TVssUsageType;
  Source : TVssSourceType;
  i,j : integer;
  pComponent : IVssWMComponent;
  cIncludeFiles,cExcludeFiles,cComponents,iMappings : UInt;
  pFileDesc : IVssWMFiledesc;
  ExcludedFile : TVssFileDescriptor;
  Component : TVssComponent;
begin
  // Get writer identity information
  OleCheck(Metadata.GetIdentity(idInstance,idWriter,bstrWriterName,Usage,Source),fWriteLineToLog,
    'TVssWriter.Initialize:Metadata.GetIdentity');
  // Get the restore method
  OleCheck(Metadata.GetRestoreMethod(RestoreMethod,bstrService,bstrUserProcedure,
    WriterRestoreConditions,RebootRequiredAfterRestore,iMappings),fWriteLineToLog,
    'TVssWriter.Initialize:Metadata.GetRestoreMethod');
  // Initialize local members
  WriterName:=OleStrToString(bstrWriterName);
  SysFreeString(bstrWriterName);
  WriterId:=GuidToString(idWriter);
  InstanceId:=GuidToString(idInstance);
  supportsRestore:=WriterRestoreConditions <> VSS_WRE_NEVER;
  // Get file counts
  OleCheck(Metadata.GetFileCounts(cIncludeFiles,cExcludeFiles,cComponents),fWriteLineToLog,
    'TVssWriter.Initialize:Metadata.GetFileCounts');
  // Get exclude files
  if cExcludeFiles>0 then for i:=0 to cExcludeFiles-1 do begin
    OleCheck(Metadata.GetExcludeFile(i,pFileDesc),fWriteLineToLog,
      'TVssWriter.Initialize:Metadata.GetExcludeFile');
    // Add this descriptor to the list of excluded files
    ExcludedFile:=TVssFileDescriptor.Create;
    ExcludedFile.OnWriteLineToLog:=fWriteLineToLog;
    ExcludedFile.Initialize(pFileDesc,VSS_FDT_EXCLUDE_FILES);
    ExcludedFiles.Add(ExcludedFile);
    end;
  // Enumerate components
  if cComponents>0 then begin
    for i:=0 to cComponents-1 do begin
      // Get component
      OleCheck(Metadata.GetComponent(i,pComponent),fWriteLineToLog,
        'TVssWriter.Initialize:Metadata.GetComponent');
      // Add this component to the list of components
      Component:=TVssComponent.Create;
      Component.OnWriteLineToLog:=fWriteLineToLog;
      Component.Initialize(WriterName,pComponent);
      Components.Add(Component);
      end;
    // Discover toplevel components
    for i:=0 to cComponents-1 do begin
      Component:=Components[i] as TVssComponent;
      Component.isTopLevel:=true;
      for j:=0 to cComponents-1 do
        if (Components[j] as TVssComponent).IsAncestorOf(Component) then Component.IsTopLevel:=false;
      end;
    end;
  end;

// Initialize from a IVssWriterComponentsExt
procedure TVssWriter.InitializeComponentsForRestore(WriterComponents : IVssWriterComponentsExt);
var
  cComponents : UInt;
  iComponent : integer;
  fComponent : TVssComponent;
  pComponent : IVssComponent;
begin
  // Erase the current list of components for this writer
  Components.clear;
  // Enumerate the components from the BC document
  OleCheck(WriterComponents.GetComponentCount(cComponents),fWriteLineToLog,
    'TVssWriter.InitializeComponentsForRestore:WriterComponents.GetComponentCount');
  // Enumerate components
  for iComponent:=0 to Components.Count-1 do begin
    // Get component
    OleCheck(WriterComponents.GetComponent(iComponent,pComponent),fWriteLineToLog,
      'TVssWriter.InitializeComponentsForRestore:WriterComponents.GetComponent');
    // Add this component to the list of components
    fComponent:=TVssComponent.Create;
    fComponent.Initialize(WriterName,pComponent);
    if assigned(fWriteLineToLog) then
      fWriteLineToLog(Format('- Found component available for restore : "%s"',[fComponent.FullPath]));
    Components.Add(fComponent);
    end;
  end;

// Print summary/detailed information about this writer
procedure TVssWriter.Print(bListDetailedInfo : boolean);
var
  i : integer;
begin
  // Print writer identity information
  if assigned(fWriteLineToLog) then begin
    fWriteLineToLog(Format(sLineBreak+'* WRITER "%s"'+sLineBreak +
      '    - WriterId   = %s'+sLineBreak+
      '    - InstanceId = %s'+sLineBreak+
      '    - Supports restore events = %s'+sLineBreak+
      '    - Writer restore conditions = %s'+sLineBreak+
      '    - Restore method = %s'+sLineBreak+
      '    - Requires reboot after restore = %s'+sLineBreak,
      [WriterName,WriterId,InstanceId,BooleanToString(supportsRestore,'Yes','No'),
       GetStringFromRestoreConditions(WriterRestoreConditions),
       GetStringFromRestoreMethod(RestoreMethod),
       BooleanToString(RebootRequiredAfterRestore,'Yes','No')]));
    // Print exclude files
    fWriteLineToLog(Format('    - Excluded files (%u):',[ExcludedFiles.Count]));
    for i:=0 to ExcludedFiles.Count-1 do (ExcludedFiles[i] as TVssFileDescriptor).Print;
    // Enumerate components
    for i:=0 to Components.Count-1 do (Components[i] as TVssComponent).Print(bListDetailedInfo);
    end;
  end;

function TVssWriter.GetStringFromRestoreMethod(eRestoreMethod
   : TVssRestoreMethodEnum) : string;
begin
  case eRestoreMethod of
    VSS_RME_UNDEFINED:              Result:='VSS_RME_UNDEFINED';
    VSS_RME_RESTORE_IF_NOT_THERE:   Result:='VSS_RME_RESTORE_IF_NOT_THERE';
    VSS_RME_RESTORE_IF_CAN_REPLACE: Result:='VSS_RME_RESTORE_IF_CAN_REPLACE';
    VSS_RME_STOP_RESTORE_START:     Result:='VSS_RME_STOP_RESTORE_START';
    VSS_RME_RESTORE_TO_ALTERNATE_LOCATION: Result:='VSS_RME_RESTORE_TO_ALTERNATE_LOCATION';
    VSS_RME_RESTORE_AT_REBOOT:      Result:='VSS_RME_RESTORE_AT_REBOOT';
    VSS_RME_CUSTOM:                 Result:='VSS_RME_CUSTOM';
    VSS_RME_RESTORE_STOP_START:     Result:='VSS_RME_RESTORE_STOP_START';
  else Result:='Undefined';
    end;
  end;

function TVssWriter.GetStringFromRestoreConditions(eRestoreEnum
   : TVssWriterRestoreEnum) : string;
begin
  case eRestoreEnum of
    VSS_WRE_UNDEFINED:        Result:='VSS_WRE_UNDEFINED';
    VSS_WRE_NEVER:            Result:='VSS_WRE_NEVER';
    VSS_WRE_IF_REPLACE_FAILS: Result:='VSS_WRE_IF_REPLACE_FAILS';
    VSS_WRE_ALWAYS:           Result:='VSS_WRE_ALWAYS';
  else Result:='Undefined';
    end;
  end;

/// /////////////////////////////////////////////////////////////////////////////////
constructor TVolumeShadowCopy.Create;
begin
  inherited Create;
  CoInitializeCalled:=false;
  VssContext:=VSS_CTX_BACKUP;
  LatestSnapshotSetID:=GUID_NULL;
  DuringRestore:=false;
  LogFileName:=SetPathDelimiter(TempDirectory)+defLogFile;
  fWriteLog:=false;
  WriterList:=TObjectList.Create;
  fShowStatus:=nil;
  end;

destructor TVolumeShadowCopy.Destroy;
begin
  if fWriteLog then CloseFile(LogFile);
  WriterList.Free;
  // Release the IVssBackupComponents interface
  // WARNING : this must be done BEFORE calling CoUninitialize()
  VssBackupComponents:=nil;
  // Call CoUninitialize if the CoInitialize was performed successfully
  if CoInitializeCalled then CoUninitialize;
  inherited Destroy;
  end;

// Convert a writer status into a string
function TVolumeShadowCopy.GetStringFromWriterStatus(eWriterStatus : TVssWriterState) : string;
begin
  case eWriterStatus of
    VSS_WS_STABLE:               Result:='Writer state : stable';
    VSS_WS_WAITING_FOR_FREEZE:   Result:='Writer state : waiting for freeze';
    VSS_WS_WAITING_FOR_THAW:     Result:='Writer state : waiting for thaw';
    VSS_WS_WAITING_FOR_POST_SNAPSHOT:  Result:='Writer state : waiting for post snapshot';
    VSS_WS_WAITING_FOR_BACKUP_COMPLETE: Result:='Writer state : waiting for backup complete';
    VSS_WS_FAILED_AT_IDENTIFY:   Result:='Writer state : failed at identy';
    VSS_WS_FAILED_AT_PREPARE_BACKUP:   Result:='Writer state : failed at prepare backup';
    VSS_WS_FAILED_AT_PREPARE_SNAPSHOT: Result:='Writer state : failed at prepare snapshot';
    VSS_WS_FAILED_AT_FREEZE:     Result:='Writer state : failed at freeze';
    VSS_WS_FAILED_AT_THAW:       Result:='Writer state : failed at thaw';
    VSS_WS_FAILED_AT_POST_SNAPSHOT:    Result:='Writer state : failed at post snapshot';
    VSS_WS_FAILED_AT_BACKUP_COMPLETE:  Result:='Writer state : failed at backup complete';
    VSS_WS_FAILED_AT_PRE_RESTORE:      Result:='Writer state : failed at restore';
    VSS_WS_FAILED_AT_POST_RESTORE:     Result:='Writer state : failed at post restore';
    else Result:=Format('Unknown constant: %d',[integer(eWriterStatus)]);
    end;
  end;

procedure TVolumeShadowCopy.WinCheck(Result : boolean; Hint : string = '');
begin
  if not Result then WinCheckError(GetLasterror,Hint);
  end;

procedure TVolumeShadowCopy.WinCheckError(LastError : integer; Hint : string = '');
begin
  if LastError<>NOERROR then begin
    if fWriteLog then begin
      if length(Hint) = 0 then Hint:='<unknown>';
      WriteLineToLog(Format(sLineBreak+'ERROR : Win32 call "%s" failed.',[Hint]));
      WriteLineToLog(Format(' - GetLastError = %d',[LastError]));
      WriteLineToLog(Format(' - Error text: %s',[SysErrorMessage(LastError)]));
      end;
    RaiseLastOSError(LastError);
    end;
  end;

procedure TVolumeShadowCopy.OleCheck(Result : HResult; Hint : string = '');
begin
  VssUtils.OleCheck(Result,WriteLineToLog,Hint);
  end;

function TVolumeShadowCopy.DisplayNameForVolume(const volumeName : string) : string;
begin
  if not GetDisplayNameForVolume(volumeName,Result) then Result:='?';
  end;

procedure TVolumeShadowCopy.SetLogFilename(const Value : string);
begin
  if fWriteLog then begin
    if not AnsiSameText(Value,fLogFileName) then begin
      CloseFile(LogFile);
      fWriteLog:=AppendFile(Value,LogFile);
      end;
    end;
  fLogFileName:=Value;
  end;

procedure TVolumeShadowCopy.SetWriteLog(Value : boolean);
begin
  if Value then begin // enable log
    if not fWriteLog then fWriteLog:=AppendFile(fLogFileName,LogFile);
    end
  else begin // disable log
    if fWriteLog then begin
      CloseFile(LogFile);
      fWriteLog:=false;
      end
    end;
  end;

procedure TVolumeShadowCopy.WriteLineToLog(const s : string);
begin
  if fWriteLog then begin
    writeln(LogFile,s);
    Flush(LogFile);
    end;
  end;

procedure TVolumeShadowCopy.ShowStatusMsg(nMsg : TVssMessages; lbs : TLineBreaks);
var
  s : string;
begin
  if assigned(fShowStatus) then fShowStatus(LoadResString(VssResourceStrings[nMsg]));
  if fWriteLog then begin
    s:=VssMessageStrings[nMsg];
    if lbLeft in lbs then s:=sLineBreak+s;
    if lbRight in lbs then s:=s+sLineBreak;
    WriteLineToLog(s);
    end;
  end;

procedure TVolumeShadowCopy.ShowFormatStatusMsg(nMsg : TVssMessages;
                  const Args: array of const; lbs : TLineBreaks);
var
  s : string;
begin
  if assigned(fShowStatus) then fShowStatus(Format(LoadResString(VssResourceStrings[nMsg]),Args));
  if fWriteLog then begin
    s:=Format(VssMessageStrings[nMsg],Args);
    if lbLeft in lbs then s:=sLineBreak+s;
    if lbRight in lbs then s:=s+sLineBreak;
    WriteLineToLog(s);
    end;
  end;

procedure TVolumeShadowCopy.PrintSnapshotProperties(Prop : TVssSnapshotProp);
var
  lAttributes : LONG;
  attributes : string;
begin
  if fWriteLog then with Prop do begin
    lAttributes:=m_lSnapshotAttributes;
    WriteLineToLog(Format('* SNAPSHOT ID = %s ...',[GuidToString(m_SnapshotId)]));
    WriteLineToLog(Format('   - Shadow copy Set: %s',[GuidToString(m_SnapshotSetId)]));
    WriteLineToLog(Format('   - Original count of shadow copies = %d',[m_lSnapshotsCount]));
    WriteLineToLog(Format('   - Original Volume name: %s [%s]',[m_pwszOriginalVolumeName,
      DisplayNameForVolume(m_pwszOriginalVolumeName)]));
    WriteLineToLog(Format('   - Creation Time: %s',[VssTimeToString(m_tsCreationTimestamp)]));
    WriteLineToLog(Format('   - Shadow copy device name: %s',[m_pwszSnapshotDeviceObject]));
    WriteLineToLog(Format('   - Originating machine: %s',[prop.m_pwszOriginatingMachine]));
    WriteLineToLog(Format('   - Service machine: %s',[prop.m_pwszServiceMachine]));
    if m_lSnapshotAttributes and VSS_VOLSNAP_ATTR_EXPOSED_LOCALLY<>0 then
        WriteLineToLog(Format('   - Exposed locally as: %s',[m_pwszExposedName]))
    else if m_lSnapshotAttributes and VSS_VOLSNAP_ATTR_EXPOSED_REMOTELY<>0 then begin
      WriteLineToLog(Format('   - Exposed remotely as %s',[m_pwszExposedName]));
      if length(m_pwszExposedPath) > 0 then
            WriteLineToLog(Format('   - Path exposed: %s',[m_pwszExposedPath]))
      else WriteLineToLog('   - Not Exposed');
      end;
    WriteLineToLog(Format('   - Provider id: %s',[GuidToString(m_ProviderId)]));
    // Display the attributes
    attributes:='';
    if lAttributes and VSS_VOLSNAP_ATTR_TRANSPORTABLE <>0 then
        attributes:=attributes+' Transportable';
    if lAttributes and VSS_VOLSNAP_ATTR_NO_AUTO_RELEASE <>0 then
        attributes:=attributes+' No_Auto_Release'
    else attributes:=attributes+' Auto_Release';
    if lAttributes and VSS_VOLSNAP_ATTR_PERSISTENT <>0 then
        attributes:=attributes+' Persistent';
    if lAttributes and VSS_VOLSNAP_ATTR_CLIENT_ACCESSIBLE <>0 then
        attributes:=attributes+' Client_accessible';
    if lAttributes and VSS_VOLSNAP_ATTR_HARDWARE_ASSISTED <>0 then
        attributes:=attributes+' Hardware';
    if lAttributes and VSS_VOLSNAP_ATTR_NO_WRITERS <>0 then
        attributes:=attributes+' No_Writers';
    if lAttributes and VSS_VOLSNAP_ATTR_IMPORTED <>0 then
        attributes:=attributes+' Imported';
    if lAttributes and VSS_VOLSNAP_ATTR_PLEX <>0 then
        attributes:=attributes+' Plex';
    if lAttributes and VSS_VOLSNAP_ATTR_DIFFERENTIAL <>0 then
        attributes:=attributes+' Differential';
    WriteLineToLog(Format('   - Attributes: %s',[attributes]));
    WriteLineToLog('');
    end;
  end;

procedure TVolumeShadowCopy.Initialize(AContext : DWord; AXmlDoc : string = '';
  ADuringRestore : boolean = false);
begin
  WriteLineToLog('');
  if not IsVssAvailable then begin
    raise EVssError.Create(rsVssNotAvail);
    end;
  ShowStatusMsg(nInitBackupComps);
  // Initialize COM
  OleCheck(CoInitialize(nil),'TVolumeShadowCopy.Initialize:CoInitialize');
  CoInitializeCalled:=true;
  // Create the internal backup components object
  OleCheck(CreateVssBackupComponents(VssBackupComponents),
    'TVolumeShadowCopy.Initialize:CreateVssBackupComponents');
  // We are during restore now?
  DuringRestore:=ADuringRestore;
  // Call either Initialize for backup or for restore
  if DuringRestore then OleCheck(VssBackupComponents.InitializeForRestore(PChar(AXmlDoc)),
    'TVolumeShadowCopy.Initialize:VssBackupComponents.InitializeForRestore')
  else begin
    // Initialize for backup
    if length(AXmlDoc) = 0 then
      OleCheck(VssBackupComponents.InitializeForBackup(nil),
      'TVolumeShadowCopy.Initialize:VssBackupComponents.InitializeForBackup')
    else OleCheck(VssBackupComponents.InitializeForBackup(PChar(AXmlDoc)),
      'TVolumeShadowCopy.Initialize:VssBackupComponents.InitializeForBackup')
    end;
  // Keep the context
  VssContext:=AContext;
  if VssContext<>VSS_CTX_BACKUP then
    OleCheck(VssBackupComponents.SetContext(VssContext),
      'TVolumeShadowCopy.Initialize:VssBackupComponents.SetContext');
  // Set various properties per backup components instance
  OleCheck(VssBackupComponents.SetBackupState(true,false,VSS_BT_COPY,false),
    'TVolumeShadowCopy.Initialize:VssBackupComponents.SetBackupState');
  end;

// Waits for the completion of the asynchronous operation
procedure TVolumeShadowCopy.WaitAndCheckForAsyncOperation(Async : IVssAsync);
var
  hrReturned : HResult;
begin
  WriteLineToLog('  => Waiting for the asynchronous operation to finish ...');
  // Wait until the async operation finishes
  OleCheck(Async.Wait(INFINITE),
    'TVolumeShadowCopy.WaitAndCheckForAsyncOperation:Async.Wait');
  // Check the result of the asynchronous operation
  hrReturned:=S_OK;
  OleCheck(Async.QueryStatus(hrReturned,nil),
    'TVolumeShadowCopy.WaitAndCheckForAsyncOperation:Async.QueryStatus');
  // Check if the async operation succeeded...
  OleCheck(hrReturned,'TVolumeShadowCopy.WaitAndCheckForAsyncOperation:Async.QueryStatus:hrReturned');
  end;

procedure TVolumeShadowCopy.InitializeWriterMetadata;
var
  cWriters : UInt;
  iWriter : integer;
  idInstance : TGuid;
  pMetadata : IVssExamineWriterMetadata;
  fWriter : TVssWriter;
begin
  ShowStatusMsg(nInitMetaData);
  // Get the list of writers in the metadata
  cWriters:=0;
  OleCheck(VssBackupComponents.GetWriterMetadataCount(cWriters),
    'TVolumeShadowCopy.InitializeWriterMetadata:VssBackupComponents.GetWriterMetadataCount');
  // Enumerate writers
  for iWriter:=0 to cWriters-1 do begin
    // Get the metadata for this particular writer
    idInstance:=GUID_NULL;
    OleCheck(VssBackupComponents.GetWriterMetadata(iWriter,idInstance,pMetadata),
      'TVolumeShadowCopy.InitializeWriterMetadata:VssBackupComponents.GetWriterMetadata');
    fWriter:=TVssWriter.Create;
    fWriter.fWriteLineToLog:=WriteLineToLog;
    fWriter.Initialize(pMetadata);
    // Add this writer to the list
    WriterList.Add(fWriter);;
    end;
  end;

// Gather writers metadata
procedure TVolumeShadowCopy.GatherWriterMetadata;
var
  Async : IVssAsync;
begin
  ShowStatusMsg(nGatherMetaData);
  // WARNING : this call can be performed only once per IVssBackupComponents instance!
  OleCheck(VssBackupComponents.GatherWriterMetadata(Async),
    'TVolumeShadowCopy.GatherWriterMetadata:VssBackupComponents.GatherWriterMetadata');
  // Waits for the async operation to finish and checks the result
  WaitAndCheckForAsyncOperation(Async);
  // Initialize the internal metadata data structures
  InitializeWriterMetadata;
  end;

function TVolumeShadowCopy.IsWriterSelected(guidInstanceId : TGuid) : boolean;
var
  i : integer;
  finstanceId : string;
begin
  // If this writer was not selected for backup,ignore it
  finstanceId:=GuidToString(guidInstanceId);
  for i:=0 to WriterList.Count-1 do with WriterList[i] as TVssWriter do
    if AnsiSameText(finstanceId,InstanceId) and not IsExcluded then begin
      Result:=true; Exit;
      end;
  Result:=false;
  end;

// Gather writers status
procedure TVolumeShadowCopy.GatherWriterStatus;
var
  pAsync : IVssAsync;
begin
  ShowStatusMsg(nGatherWriterStatus);
  // WARNING : GatherWriterMetadata must be called before
  OleCheck(VssBackupComponents.GatherWriterStatus(pAsync),
    'TVolumeShadowCopy.GatherWriterStatus:ssBackupComponents.GatherWriterStatus');
  // Waits for the async operation to finish and checks the result
  WaitAndCheckForAsyncOperation(pAsync);
  end;

// Check the status for all selected writers
procedure TVolumeShadowCopy.CheckSelectedWriterStatus;
var
  cWriters : UInt;
  i : integer;
  idInstance,idWriter : TGuid;
  eWriterStatus : TVssWriterState;
  bstrWriterName : PWideChar;
  hrWriterFailure : HResult;
  fWriterName : string;
begin
  if (VssContext and VSS_VOLSNAP_ATTR_NO_WRITERS) = 0 then begin
    // Gather writer status to detect potential errors
    GatherWriterStatus;
    // Gets the number of writers in the gathered status info
    // (WARNING : GatherWriterStatus must be called before)
    OleCheck(VssBackupComponents.GetWriterStatusCount(cWriters),
      'TVolumeShadowCopy.CheckSelectedWriterStatus:VssBackupComponents.GetWriterStatusCount');
    // Enumerate each writer
    if cWriters>0 then for i:=0 to cWriters-1 do begin
      // Get writer status
      OleCheck(VssBackupComponents.GetWriterStatus(i,idInstance,idWriter,
        bstrWriterName,eWriterStatus,hrWriterFailure),
        'TVolumeShadowCopy.CheckSelectedWriterStatus:VssBackupComponents.GetWriterStatus');
      fWriterName:=OleStrToString(bstrWriterName);
      SysFreeString(bstrWriterName);
      // If the writer is not selected, just continue
      if not IsWriterSelected(idInstance) then continue;
      // If the writer is in non-stable state, break
      case eWriterStatus of
        VSS_WS_FAILED_AT_IDENTIFY,VSS_WS_FAILED_AT_PREPARE_BACKUP,
        VSS_WS_FAILED_AT_PREPARE_SNAPSHOT,VSS_WS_FAILED_AT_FREEZE,
        VSS_WS_FAILED_AT_THAW,VSS_WS_FAILED_AT_POST_SNAPSHOT,
        VSS_WS_FAILED_AT_BACKUP_COMPLETE,VSS_WS_FAILED_AT_PRE_RESTORE,
        VSS_WS_FAILED_AT_POST_RESTORE: begin
            if fWriteLog then
              WriteLineToLog(Format('ERROR : Selected writer "%s" is in failed state!'+sLineBreak+
                'Status: %d (%s)'+sLineBreak+'Writer Failure code : $%.8x'+sLineBreak+
                'Writer ID : {%s}'+sLineBreak+'Instance ID : {%s}',
                [fWriterName,integer(eWriterStatus),GetStringFromWriterStatus(eWriterStatus),
                 hrWriterFailure,GuidToString(idWriter),GuidToString(idInstance)]));
            raise EOleSysError.Create('',integer(eWriterStatus),0);
            end;
        end;
      end;
    end;
  end;

// Lists the writer metadata
procedure TVolumeShadowCopy.ListWriterMetadata(bListDetailedInfo : boolean);
var
  iWriter : integer;
begin
  WriteLineToLog('Listing writer metadata ...');
    // Enumerate writers
  for iWriter:=0 to WriterList.Count-1 do
    (WriterList[iWriter] as TVssWriter).Print(bListDetailedInfo);
  end;

// Lists the status for all writers
procedure TVolumeShadowCopy.ListWriterStatus;
var
  cWriters : UInt;
  iWriter : integer;
  idInstance,idWriter : TGuid;
  eWriterStatus : TVssWriterState;
  bstrWriterName : PWideChar;
  hrWriterFailure : HResult;
begin
  WriteLineToLog('Listing writer status ...');
  // Gets the number of writers in the gathered status info
  // (WARNING: GatherWriterStatus must be called before)
  OleCheck(VssBackupComponents.GetWriterStatusCount(cWriters));
  WriteLineToLog(Format('- Number of writers that responded: %u',[cWriters]));
  // Enumerate each writer
  for iWriter:=0 to WriterList.Count-1 do begin
    // Get writer status
    OleCheck(VssBackupComponents.GetWriterStatus(iWriter,idInstance,idWriter,
                             bstrWriterName,eWriterStatus,hrWriterFailure));
    // Print writer status
    if fWriteLog then WriteLineToLog(Format(sLineBreak+'* WRITER "%s"'+sLineBreak
        +'   - Status: %d (%s)'+sLineBreak
        +'   - Writer Failure code: $%.8x (%s)'+sLineBreak
        +'   - Writer ID: %s'+sLineBreak
        +'   - Instance ID: %s',
        [bstrWriterName,integer(eWriterStatus),GetStringFromWriterStatus(eWriterStatus),
        hrWriterFailure,HResultToString(hrWriterFailure),
        GuidToString(idWriter),GuidToString(idInstance)]));
    end;
  end;

// Discover directly excluded components (that were excluded through the command-line)
procedure TVolumeShadowCopy.DiscoverDirectlyExcludedComponents(ExcludedWriterAndComponentList : TStringList;
                                                               AWriterList : TObjectList);
var
  iWriter,iComponent : integer;
  fWriter : TVssWriter;
  fComponent : TVssComponent;
  nonExcludedComponents,excluded : boolean;
  componentPathWithWriterName,componentPathWithWriterID,
    componentPathWithWriterIID : string;
begin
  ShowStatusMsg(nDscvDirExclComps);
  // Discover components that should be excluded from the shadow set
  // This means components that have at least one File Descriptor requiring
  // volumes not in the shadow set.
  for iWriter:=0 to AWriterList.Count-1 do begin
    fWriter:=AWriterList[iWriter] as TVssWriter;
    // Check if the writer is excluded
    if assigned(ExcludedWriterAndComponentList) then with ExcludedWriterAndComponentList do
      excluded:=(IndexOf(fWriter.WriterName)>=0) or
        (IndexOf(fWriter.WriterId)>=0) or (IndexOf(fWriter.InstanceId)>=0)
    else excluded:=false;
    if excluded then fWriter.IsExcluded:=true
    else begin
      // Check if the component is excluded
      for iComponent:=0 to fWriter.Components.Count-1 do begin
        fComponent:=fWriter.Components[iComponent] as TVssComponent;
        // Check to see if this component is explicitly excluded
        // Compute various component paths
        // Format : Writer:logicaPath\componentName
        componentPathWithWriterName:=fWriter.WriterName+':'+fComponent.FullPath;
        componentPathWithWriterID:=fWriter.WriterId+':'+fComponent.FullPath;
        componentPathWithWriterIID:=fWriter.InstanceId+':'+fComponent.FullPath;
        // Check to see if this component is explicitly excluded
        if assigned(ExcludedWriterAndComponentList) then with ExcludedWriterAndComponentList do
          excluded:=(IndexOf(componentPathWithWriterName)>=0) or
            (IndexOf(componentPathWithWriterID)>=0) or
            (IndexOf(componentPathWithWriterIID)>=0)
        else excluded:=false;
        if excluded then begin
          WriteLineToLog (Format('- Component "%s" from writer "%s" is explicitly excluded from backup',
            [fComponent.FullPath,fWriter.WriterName]));
          fComponent.IsExcluded:=true;
          end
        end;
      // Now, discover if we have any selected components. If none, exclude the whole writer
      nonExcludedComponents:=false;
      for iComponent:=0 to fWriter.Components.Count-1 do begin
        fComponent:=fWriter.Components[iComponent] as TVssComponent;
        if not fComponent.IsExcluded then nonExcludedComponents:=true;
        end;
      // If all components are missing or excluded, then exclude the writer too
      if not nonExcludedComponents then begin
        WriteLineToLog(Format('- Excluding writer "%s" since it has no selected components for restore.',
          [fWriter.WriterName]));
        fWriter.IsExcluded:=true;
        end;
      end;
    end;
  end;

// Discover excluded components that have file groups outside the shadow set
procedure TVolumeShadowCopy.DiscoverNonShadowedExcludedComponents(ShadowSourceVolumes : TStringList);
var
  iWriter,iComponent,iVol : integer;
  fWriter : TVssWriter;
  fComponent : TVssComponent;
  wsUniquePath : array [0 .. MAX_PATH] of WideChar;
  wsLocalVolume : string;
begin
  ShowStatusMsg(nDscvNsExclComps);
  // Discover components that should be excluded from the shadow set
  // This means components that have at least one File Descriptor requiring
  // volumes not in the shadow set.
  for iWriter:=0 to WriterList.Count-1 do begin
    fWriter:=WriterList[iWriter] as TVssWriter;
    // Check if the writer is excluded
    WriteLineToLog(Format(' * Writer: "%s":',[fWriter.WriterName]));
    if not fWriter.IsExcluded then begin
      // Check if the component is excluded
      for iComponent:=0 to fWriter.Components.Count-1 do begin
        fComponent:=fWriter.Components[iComponent] as TVssComponent;
        // Check to see if this component is explicitly excluded
        WriteLineToLog(Format('  * Component: "%s":',[fComponent.FullPath]));
        if not fComponent.IsExcluded then begin
          // Try to find an affected volume outside the shadow set
          // If yes,exclude the component
          for iVol:=0 to fComponent.AffectedVolumes.Count-1 do begin
            if (ClusterIsPathOnSharedVolume(PWideChar(fComponent.AffectedVolumes[iVol]))) then begin
              ClusterGetVolumeNameForVolumeMountPoint(PWideChar(fComponent.AffectedVolumes[iVol]),
              wsUniquePath,MAX_PATH+1);
              fComponent.AffectedVolumes[iVol]:=wsUniquePath;
              end;
            if ShadowSourceVolumes.IndexOf(fComponent.AffectedVolumes[iVol])<0 then begin
              if (GetDisplayNameForVolume(fComponent.AffectedVolumes[iVol],wsLocalVolume)) then
                WriteLineToLog(Format('    Component "%s" from writer "%s" is excluded from backup (it requires "%s" in the shadow set)',
                  [fComponent.FullPath,fWriter.WriterName,wsLocalVolume]))
              else WriteLineToLog(Format('    Component "%s" from writer "%s" is excluded from backup',
                  [fComponent.FullPath,fWriter.WriterName]));
              fComponent.IsExcluded:=true;
              Break;
              end;
            end;
          end
        else WriteLineToLog('  - Excluded');
        end;
      end
    else WriteLineToLog('  - Excluded');
    end;
  end;

// Discover the components that should not be included (explicitly or implicitly)
// These are componenets that are have directly excluded descendents
procedure TVolumeShadowCopy.DiscoverAllExcludedComponents;
var
  iWriter,iComponent,i : integer;
  fWriter : TVssWriter;
  fComponent,descendent : TVssComponent;
begin
  ShowStatusMsg(nDscvAllExclComps);
  // Discover components that should be excluded from the shadow set
  // This means components that have at least one File Descriptor requiring
  // volumes not in the shadow set.
  for iWriter:=0 to WriterList.Count-1 do begin
    fWriter:=WriterList[iWriter] as TVssWriter;
    if not fWriter.IsExcluded then begin // Ignore explicitly excluded writers
      // Enumerate all components
      for iComponent:=0 to fWriter.Components.Count-1 do begin
        fComponent:=fWriter.Components[iComponent] as TVssComponent;
        // Check if this component has any excluded children
        // If yes, deselect it
        for i:=0 to fWriter.Components.Count-1 do begin
          descendent:=fWriter.Components[i] as TVssComponent;
          if fComponent.IsAncestorOf(descendent) and descendent.IsExcluded then begin
            WriteLineToLog(Format('- Component "%s" from writer "%s" is excluded from backup '
             +'(it has an excluded descendent: %s)',
             [fComponent.FullPath,fWriter.WriterName,descendent.CompName]));
            fComponent.IsExcluded:=true;
            Break;
            end;
          end;
        end;
      end;
    end;
  end;

// Discover excluded writers. These are writers that:
//-either have a top-level nonselectable excluded component
//-or do not have any included components (all its components are excluded)
procedure TVolumeShadowCopy.DiscoverExcludedWriters;
var
  iWriter,iComponent : integer;
  fWriter : TVssWriter;
  fComponent : TVssComponent;
begin
  ShowStatusMsg(nDscvExclWriters);
  // Enumerate writers
  for iWriter:=0 to WriterList.Count-1 do
  begin
    fWriter:=WriterList[iWriter] as TVssWriter;
    if not fWriter.IsExcluded then begin // Ignore explicitly excluded writers
      // Discover if we have any:
      //-non-excluded selectable components
      //-or non-excluded top-level non-selectable components
      // If we have none, then the whole writer must be excluded from the backup
      fWriter.IsExcluded:=true;
      for iComponent:=0 to fWriter.Components.Count-1 do begin
        fComponent:=fWriter.Components[iComponent] as TVssComponent;
        if fComponent.CanBeExplicitlyIncluded then begin
          fWriter.IsExcluded:=false;
          Break;
          end;
        end;
      // No included components were found
      if fWriter.IsExcluded then begin
        WriteLineToLog(Format('- The writer "%s" is now entirely excluded from the backup:',
          [fWriter.WriterName]));
        WriteLineToLog('  (it does not contain any components that can be potentially included in the backup)');
        end
      else begin
        // Now, discover if we have any top-level excluded non-selectable component
        // If this is true, then the whole writer must be excluded from the backup
        for iComponent:=0 to fWriter.Components.Count-1 do begin
          fComponent:=fWriter.Components[iComponent] as TVssComponent;
          if fComponent.IsTopLevel and not fComponent.IsSelectable and fComponent.IsExcluded then begin
            WriteLineToLog(Format('- The writer "%s" is now entirely excluded from the backup:',
              [fWriter.WriterName]));
            WriteLineToLog(Format('  (the top-level non-selectable component "%s" is an excluded component)',
              [fComponent.FullPath]));
            fWriter.IsExcluded:=true;
            Break;
            end;
          end;
        end;
      end;
    end;
  end;

// Discover the components that should be explicitly included
// These are any included top components
procedure TVolumeShadowCopy.DiscoverExplicitelyIncludedComponents;
var
  iWriter,iComponent,i : integer;
  fWriter : TVssWriter;
  fComponent,ancestor : TVssComponent;
begin
  ShowStatusMsg(nDscvExpInclComps);
  // Enumerate all writers
  for iWriter:=0 to WriterList.Count-1 do begin
    fWriter:=WriterList[iWriter] as TVssWriter;
    if not fWriter.IsExcluded then begin // Ignore explicitly excluded writers
      // Compute the roots of included components
      for iComponent:=0 to fWriter.Components.Count-1 do begin
        fComponent:=fWriter.Components[iComponent] as TVssComponent;
        if fComponent.CanBeExplicitlyIncluded then begin
          // Test if our component has a parent that is also included
          fComponent.IsExplicitlyIncluded:=true;
          for i:=0 to fWriter.Components.Count-1 do begin
            ancestor:=fWriter.Components[i] as TVssComponent;
            if ancestor.IsAncestorOf(fComponent) and ancestor.CanBeExplicitlyIncluded then begin
              // This cannot be explicitely included since we have another
              // ancestor that that must be (implictely or explicitely) included
              fComponent.IsExplicitlyIncluded:=false;
              Break;
              end;
            end;
          end;
        end;
      end;
    end;
  end;

// Verify that the given components will be explicitly or implicitly selected
procedure TVolumeShadowCopy.VerifyExplicitelyIncludedComponent(const AIncludedComponent : string;
                 AWriterList : TObjectList);
var
  iWriter,iComponent,k : integer;
  fWriter : TVssWriter;
  fComponent,ancestor : TVssComponent;
  componentPathWithWriterName,componentPathWithWriterID,
  componentPathWithWriterIID : string;
  isIncluded : boolean;
begin
  WriteLineToLog(Format('- Verifing component "%s" ...',[AIncludedComponent]));
  // Enumerate all writers
  for iWriter:=0 to WriterList.Count-1 do begin
    fWriter:=WriterList[iWriter] as TVssWriter;
    if not fWriter.IsExcluded then begin // Ignore explicitly excluded writers
      // Find the associated component
      for iComponent:=0 to fWriter.Components.Count-1 do begin
        fComponent:=fWriter.Components[iComponent] as TVssComponent;
        // Ignore explicitly excluded components
        if not fComponent.IsExcluded then begin
          // Compute various component paths
          // Format : Writer:logicaPath\componentName
          componentPathWithWriterName:=fWriter.WriterName+':'+fComponent.FullPath;
          componentPathWithWriterID:=fWriter.WriterId+':'+fComponent.FullPath;
          componentPathWithWriterIID:=fWriter.InstanceId+':'+fComponent.FullPath;
          // Check to see if this component is (implicitly or explicitly) included
          if (AnsiSameText(componentPathWithWriterName,AIncludedComponent) or
              AnsiSameText(componentPathWithWriterID,AIncludedComponent) or
              AnsiSameText(componentPathWithWriterIID,AIncludedComponent)) then begin
            // If we are during restore, we just found our component
            if DuringRestore then begin
              WriteLineToLog(Format(' - The component "%s" is selected',[AIncludedComponent]));
              Exit;
              end;
            // If not explicitly included, check to see if there is an explicitly included ancestor
            isIncluded:=fComponent.IsExplicitlyIncluded;
            if not isIncluded then begin
              for k:=0 to fWriter.Components.Count-1 do begin
                ancestor:=fWriter.Components[k] as TVssComponent;
                if ancestor.IsAncestorOf(fComponent) and ancestor.IsExplicitlyIncluded then begin
                  isIncluded:=true;
                  Break;
                  end;
                end;
              end;
            if isIncluded then begin
              WriteLineToLog(Format(' - The component "%s" is selected',[AIncludedComponent]));
              Exit;
              end
            else begin
              WriteLineToLog(Format('ERROR : The component "%s" was not included in the backup! Aborting backup ...',
                [AIncludedComponent]));
              WriteLineToLog('- Please reveiw the component/subcomponent definitions');
              WriteLineToLog('- Also, please verify list of volumes to be shadow copied.');
              raise EOleSysError.Create('',E_INVALIDARG,0);
              end;
            end;
          end;
        end;
      end;
    end;
  WriteLineToLog(Format('ERROR : The component "%s" was not found in the writer components list! Aborting backup ...',
    [AIncludedComponent]));
  WriteLineToLog('- Please check the syntax of the component name.');
  raise EOleSysError.Create('',E_INVALIDARG,0);
  end;

// Verify that all the components of this writer are selected
procedure TVolumeShadowCopy.VerifyExplicitelyIncludedWriter(const AWriterName : string; AWriterList : TObjectList);
var
  iWriter,iComponent : integer;
  fWriter : TVssWriter;
  fComponent : TVssComponent;
begin
  WriteLineToLog(Format('- Verifing that all components of writer "%s" are included in backup ...',
    [AWriterName]));
  // Enumerate all writers
  for iWriter:=0 to WriterList.Count-1 do begin
    fWriter:=WriterList[iWriter] as TVssWriter;
    if not fWriter.IsExcluded then begin // Ignore explicitly excluded writers
      // Check if we found the writer
      if AnsiSameText(AWriterName,fWriter.WriterName) or
        AnsiSameText(AWriterName,fWriter.WriterId) or
        AnsiSameText(AWriterName,fWriter.InstanceId) then begin
        if fWriter.IsExcluded then begin
          WriteLineToLog(Format('ERROR : The writer "%s" was not included in the backup! Aborting backup ...',
            [AWriterName]));
          WriteLineToLog('- Please review the component/subcomponent definitions');
          WriteLineToLog('- Also, please verify list of volumes to be shadow copied.');
          raise EOleSysError.Create('',E_INVALIDARG,0);
          end;
        // Make sure all its associated components are selected
        for iComponent:=0 to fWriter.Components.Count-1 do begin
          fComponent:=fWriter.Components[iComponent] as TVssComponent;
          if fComponent.IsExcluded then begin
            if fWriteLog then begin
              WriteLineToLog(Format('ERROR : The writer "%s" has components not included in the backup! Aborting backup ...',
                [AWriterName]));
              WriteLineToLog(Format('- The component "%s" was not included in the backup.',
                [fComponent.FullPath]));
              WriteLineToLog('- Please reveiw the component/subcomponent definitions');
              WriteLineToLog('- Also, please verify list of volumes to be shadow copied.');
              end;
            raise EOleSysError.Create('',E_INVALIDARG,0);
            end;
          end;
        WriteLineToLog(Format('  - All components from writer "%s" are selected',[AWriterName]));
        Exit;
        end;
      end;
    end;
  WriteLineToLog(Format('"ERROR : The writer \"%s\" was not found! Aborting backup ...',[AWriterName]));
  WriteLineToLog('- Please check the syntax of the writer name/id.');
  raise EOleSysError.Create('',E_INVALIDARG,0);
  end;

// Discover the components that should be explicitly included
// These are any included top components
procedure TVolumeShadowCopy.SelectExplicitelyIncludedComponents;
var
  iWriter,iComponent : integer;
  fWriter : TVssWriter;
  fComponent : TVssComponent;
begin
  ShowStatusMsg(nSelExpInclComps);
  // Enumerate all writers
  for iWriter:=0 to WriterList.Count-1 do begin
    fWriter:=WriterList[iWriter] as TVssWriter;
    if not fWriter.IsExcluded then begin // Ignore explicitly excluded writers
      WriteLineToLog(Format(' * Writer "%s":',[fWriter.WriterName]));
      // Compute the roots of included components
      for iComponent:=0 to fWriter.Components.Count-1 do begin
        fComponent:=fWriter.Components[iComponent] as TVssComponent;
        if fComponent.IsExplicitlyIncluded then begin
          WriteLineToLog(Format('  - Add component %s',[fComponent.FullPath]));
          // Add the component
          OleCheck(VssBackupComponents.AddComponent(StringToGuid(fWriter.InstanceId),
            StringToGuid(fWriter.WriterId),fComponent.CompType,PWideChar(fComponent.LogicalPath),
            PWideChar(fComponent.CompName)),
            'TVolumeShadowCopy.SelectExplicitelyIncludedComponents:VssBackupComponents.AddComponent');
          end;
        end;
      end;
    end;
  end;

// Select the maximum number of components such that their
// file descriptors are pointing only to volumes to be shadow copied
procedure TVolumeShadowCopy.SelectComponentsForBackup(ShadowSourceVolumes,
                ExcludedWriterAndComponentList,IncludedWriterAndComponentList : TStringList);
var
  i : integer;
begin
  // First, exclude all components that have data outside of the shadow set
  DiscoverDirectlyExcludedComponents(ExcludedWriterAndComponentList,WriterList);
  // Then discover excluded components that have file groups outside the shadow set
  DiscoverNonShadowedExcludedComponents(ShadowSourceVolumes);
  // Now, exclude all components that are have directly excluded descendents
  DiscoverAllExcludedComponents;
  // Next, exclude all writers that:
  //-either have a top-level nonselectable excluded component
  //-or do not have any included components (all its components are excluded)
  DiscoverExcludedWriters;
  // Now, discover the components that should be included (explicitly or implicitly)
  // These are the top components that do not have any excluded children
  DiscoverExplicitelyIncludedComponents();
  // Verify if the specified writers/components were included
  WriteLineToLog('Verifying explicitly specified writers/components ...');
  if assigned(includedWriterAndComponentList) then begin
    for i:=0 to includedWriterAndComponentList.Count-1 do begin
      // Check whether a component or a writer is specified
      if Pos(':',includedWriterAndComponentList[i]) <> 0 then
        VerifyExplicitelyIncludedComponent(IncludedWriterAndComponentList[i],WriterList)
      else
        VerifyExplicitelyIncludedWriter(IncludedWriterAndComponentList[i],WriterList);
      end;
    end;
  // Finally, select the explicitly included components
  SelectExplicitelyIncludedComponents;
  end;

// Save the backup components document
procedure TVolumeShadowCopy.SaveBackupComponentsDocument(const AFileName : string);
var
  bstrXML : widestring;
begin
  ShowStatusMsg(nSaveBackCompsDoc);
//  SetLength(bstrXML,XmlSize);
  // Get the Backup Components in XML format
  OleCheck(VssBackupComponents.SaveAsXML(@bstrXML),
    'TVolumeShadowCopy.SaveBackupComponentsDocument:VssBackupComponents.SaveAsXML');
  // Save the XML string to the file
  WriteStringToFile(AFileName,bstrXML);
  end;

// Add volumes to the shadow set
procedure TVolumeShadowCopy.AddToSnapshotSet(VolumeList : TStringList);
var
  i : integer;
  volume,s : string;
  SnapshotID : TGuid;
begin
  ShowStatusMsg(nAddVolSnapSet);
  // Preserve the list of volumes for script generation
  LatestVolumeList:=VolumeList; // never used ???
  Assert(length(LatestSnapshotIdList)=0,'ASSERTION FAILED : No empty list of latest snapshots');
  // Add volumes to the shadow set
  for i:=0 to VolumeList.Count-1 do begin
    volume:=VolumeList[i];
    if not GetDisplayNameForVolume(volume,s) then s:='??';
    WriteLineToLog(Format('- Adding volume %s [%s] to the shadow set...',[volume,s]));
    OleCheck(VssBackupComponents.AddToSnapshotSet(PWideChar(volume),GUID_NULL,SnapshotID),
      'TVolumeShadowCopy.AddToSnapshotSet:VssBackupComponents.AddToSnapshotSet');
    // Preserve this shadow ID for script generation
    AddGuidToList(LatestSnapshotIdList,SnapshotID);
    end;
  end;

// Prepare the shadow for backup
procedure TVolumeShadowCopy.PrepareForBackup;
var
  pAsync : IVssAsync;
begin
  ShowStatusMsg(nPrepBackup);
  OleCheck(VssBackupComponents.PrepareForBackup(pAsync),
    'TVolumeShadowCopy.PrepareForBackup:VssBackupComponents.PrepareForBackup');
  // Waits for the async operation to finish and checks the result
  WaitAndCheckForAsyncOperation(pAsync);
  // Check selected writer status
  CheckSelectedWriterStatus;
  end;

// Effectively creating the shadow (calling DoSnapshotSet)
procedure TVolumeShadowCopy.DoSnapshotSet;
var
  pAsync : IVssAsync;
begin
  ShowStatusMsg(nCreateShadowCopy);
  OleCheck(VssBackupComponents.DoSnapshotSet(pAsync),
    'TVolumeShadowCopy.DoSnapshotSet:VssBackupComponents.DoSnapshotSet');
  // Waits for the async operation to finish and checks the result
  WaitAndCheckForAsyncOperation(pAsync);
  // Do not attempt to continue with delayed snapshot ...
  if VssContext and VSS_VOLSNAP_ATTR_DELAYED_POSTSNAPSHOT <> 0 then begin
    WriteLineToLog(sLineBreak+'Fast DoSnapshotSet finished.'+sLineBreak);
    Exit;
    end;
  // Check selected writer status
  CheckSelectedWriterStatus;
  ShowStatusMsg(nCreateCopySuccess);
  WriteLineToLog('');
  end;

// Query all the shadow copies in the given set
// If snapshotSetID is NULL, just query all shadow copies in the system
procedure TVolumeShadowCopy.QuerySnapshotSet(SnapshotSetID : TGuid);
var
  pIEnumSnapshots : IVssEnumObject;
  hr : HResult;
  Prop : TVssObjectProp;
  ulFetched : ULONG;
begin
  if SnapshotSetID = GUID_NULL then ShowStatusMsg(nQueryShadowCopies,[lbRight])
  else ShowFormatStatusMsg(nQuerySnapshotSetID,[GuidToString(SnapshotSetID)],[lbRight]);
  // Get list of all shadow copies.
  hr:=VssBackupComponents.Query(GUID_NULL,VSS_OBJECT_NONE,VSS_OBJECT_SNAPSHOT,pIEnumSnapshots);
  OleCheck(hr,'TVolumeShadowCopy.QuerySnapshotSet:VssBackupComponents.Query');
  // If there are no shadow copies, just return
  if hr=S_FALSE then begin
    if SnapshotSetID = GUID_NULL then begin
      WriteLineToLog(sLineBreak+'There are no shadow copies in the system'+sLineBreak);
      Exit;
      end;
    end;
  try
    // Enumerate all shadow copies.
    while true do begin
      // Get the next element
      hr:=pIEnumSnapshots.Next(1,Prop,ulFetched);
      OleCheck(hr,'TVolumeShadowCopy.QuerySnapshotSet:pIEnumSnapshots.Next');
      // We reached the end of list
      if ulFetched = 0 then Break;
      // Print the shadow copy (if not filtered out)
      if (SnapshotSetID = GUID_NULL) or (Prop.PObj.Snap.m_SnapshotSetId=SnapshotSetID)
      then PrintSnapshotProperties(Prop.PObj.Snap);
      end;
  finally VssFreeSnapshotProperties(@Prop.PObj.Snap);
    end;
  end;

// Query the properties of the given shadow copy
procedure TVolumeShadowCopy.GetSnapshotProperties(SnapshotSetID : TGuid);
var
  hr : HResult;
  Snap : TVssSnapshotProp;
begin
  try
    // Get the shadow copy properties
    WriteLineToLog(Format('- Properties of shadow copy "%s" ...',[GuidToString(snapshotSetID)]));
    hr:=VssBackupComponents.GetSnapshotProperties(SnapshotSetID,Snap);
    if failed(hr) then begin
      WriteLineToLog('Error while retrieving properties of shadow copy ...');
      OleCheck(hr,'TVolumeShadowCopy.etSnapshotProperties:VssBackupComponents.GetSnapshotProperties');
      end
    else
    // Print the properties of this shadow copy
      PrintSnapshotProperties(Snap);
  finally VssFreeSnapshotProperties(@Snap);
    end;
  end;

// Create the shadow copy set
procedure TVolumeShadowCopy.CreateSnapshotSet(AVolumeList : TStringList;
                     AExcludedWriterList : TStringList = nil;
                     AIncludedWriterList : TStringList = nil; AOutXmlFile : string = '');
var
  SnapshotWithWriters : boolean;
begin
  SnapshotWithWriters:=(VssContext and VSS_VOLSNAP_ATTR_NO_WRITERS) = 0;
  if SnapshotWithWriters then begin
    // Gather writer metadata
    GatherWriterMetadata;
    // Select writer components based on the given shadow volume list
    SelectComponentsForBackup(AVolumeList,AExcludedWriterList,AIncludedWriterList);
    // Start the shadow set
    OleCheck(VssBackupComponents.StartSnapshotSet(LatestSnapshotSetID),
      'TVolumeShadowCopy.CreateSnapshotSet:VssBackupComponents.StartSnapshotSet');
    ShowFormatStatusMsg(nCreateShadowSet,[GuidToString(LatestSnapshotSetID)]);
    // Add the specified volumes to the shadow set
    AddToSnapshotSet(AVolumeList);
    // Prepare for backup.
    // This will internally create the backup components document with the selected components
    if SnapshotWithWriters then PrepareForBackup;
    // Creates the shadow set
    DoSnapshotSet;
    // Do not attempt to continue with delayed snapshot ...
    if VssContext and VSS_VOLSNAP_ATTR_DELAYED_POSTSNAPSHOT <> 0 then begin
      WriteLineToLog(sLineBreak+'Fast snapshot created. Exiting... '+sLineBreak);
      Exit;
      end;
    // Saves the backup components document, if needed
    if length(AOutXmlFile) > 0 then SaveBackupComponentsDocument(AOutXmlFile);
    // List all the created shadow copies
    if VssContext and VSS_VOLSNAP_ATTR_TRANSPORTABLE = 0 then begin
      WriteLineToLog(sLineBreak+'List of created shadow copies:');
      QuerySnapshotSet(LatestSnapshotSetID);
      end;
    end;
  end;

procedure TVolumeShadowCopy.CreateSnapshotSet(const AVolume : string);
var
  AVolumeList : TStringList;
begin
  AVolumeList:=TStringList.Create;
  AVolumeList.Add(AVolume);
  CreateSnapshotSet(AVolumeList);
  AVolumeList.Free;
  end;

// Marks all selected components as succeeded for backup
procedure TVolumeShadowCopy.SetBackupSucceeded(succeeded : boolean);
var
  iWriter,iComponent : integer;
  fWriter : TVssWriter;
  fComponent : TVssComponent;
begin
  // Enumerate writers
  for iWriter:=0 to WriterList.Count-1 do begin
    fWriter:=WriterList[iWriter] as TVssWriter;
    // Enumerate components
    for iComponent:=0 to fWriter.Components.Count-1 do begin
      fComponent:=fWriter.Components[iComponent] as TVssComponent;
      // Test that the component is explicitely selected and requires notification
      if fComponent.IsExplicitlyIncluded then
        // Call SetBackupSucceeded for this component
        OleCheck(VssBackupComponents.SetBackupSucceeded
          (StringToGuid(fWriter.InstanceId),StringToGuid(fWriter.WriterId),
          fComponent.CompType,PWideChar(fComponent.LogicalPath),
          PWideChar(fComponent.CompName),succeeded),
          'TVolumeShadowCopy.SetBackupSucceeded:VssBackupComponents.SetBackupSucceeded');
      end;
    end;
  end;

// Ending the backup (calling BackupComplete)
procedure TVolumeShadowCopy.BackupComplete(succeeded : boolean);
var
  cWriters : UInt;
  pAsync : IVssAsync;
begin
  OleCheck(VssBackupComponents.GetWriterComponentsCount(cWriters),
    'TVolumeShadowCopy.BackupComplete:VssBackupComponents.GetWriterComponentsCount');
  if cWriters=0 then WriteLineToLog('- There were no writer components in this backup')
  else begin
    if succeeded then
      WriteLineToLog('- Mark all writers as successfully backed up ... ')
    else
      WriteLineToLog('- Backup failed. Mark all writers as not successfully backed up ... ');
    SetBackupSucceeded(succeeded);
    ShowStatusMsg(nCompleteBackup,[lbLeft]);
    OleCheck(VssBackupComponents.BackupComplete(pAsync),
      'TVolumeShadowCopy.BackupComplete:VssBackupComponents.BackupComplete');
    // Waits for the async operation to finish and checks the result
    WaitAndCheckForAsyncOperation(pAsync);
    // Check selected writer status
    CheckSelectedWriterStatus;
    end;
  end;

// This is useful for management operations
function TVolumeShadowCopy.GetLastShadowDeviceName : string;
var
  i : integer;
  Snap : TVssSnapshotProp;
begin
  Result:='';
  try
    // For each added volume add the VSCSC.EXE exposure command
    for i:=0 to High(LatestSnapshotIdList) do begin
      // Get shadow copy device (if the snapshot is there)
      if VssContext and VSS_VOLSNAP_ATTR_TRANSPORTABLE = 0 then begin
        OleCheck(VssBackupComponents.GetSnapshotProperties(LatestSnapshotIdList[i],Snap),
          'TVolumeShadowCopy.GetLastShadowDeviceName:VssBackupComponents.GetSnapshotProperties');
        Result:=Snap.m_pwszSnapshotDeviceObject;
        Exit;
        end;
      end;
  finally VssFreeSnapshotProperties(@Snap);
    end;
  end;

procedure TVolumeShadowCopy.ExecuteCommand(const Command,Param : string);
var
  s : string;
  si : TStartupInfo;
  pi : TProcessInformation;
  dwExitCode : DWord;
begin
  WriteLineToLog('-----------------------------------------------------');
  WriteLineToLog(Format('- Executing command "%s %s" ...',[Command,Param]));
  // Create process to start Program
  FillChar(si,sizeof(TStartupInfo),0);
  si.cb:=sizeof(TStartupInfo);
  //
  // Security Remarks-CreateProcess
  //
  // The first parameter, lpApplicationName, can be NULL, in which case the
  // executable name must be in the white space-delimited string pointed to
  // by lpCommandLine. If the executable or path name has a space in it, there
  // is a risk that a different executable could be run because of the way the
  // function parses spaces. The following example is dangerous because the
  // function will attempt to run "Program.exe", if it exists, instead of "MyApp.exe".
  //
  // CreateProcess(NULL, "C:\\Program Files\\MyApp", ...)
  //
  // If a malicious user were to create an application called "Program.exe"
  // on a system, any program that incorrectly calls CreateProcess
  // using the Program Files directory will run this application
  // instead of the intended application.
  //
  // For this reason we blocked parameters in the executed command.
  //
  s:='"'+Command+'"'+' '+Param;
  try
    WinCheck(CreateProcess(nil,   // command line for application to exedute
      PChar(s),nil,               // Security
      nil,                        // Security
      false,                      // use InheritHandles
      NORMAL_PRIORITY_CLASS,      // Priority
      nil,                        // Environment
      nil,                        // directory
      si, pi),'CreateProcess('+s+')');
    // Wait until child process exits.
    WinCheck(WaitForSingleObject(pi.hProcess,INFINITE)=WAIT_OBJECT_0,'WaitForSingleObject');
    WriteLineToLog('- Command completed.');
    // Checking the exit code
    WinCheck(GetExitCodeProcess(pi.hProcess,dwExitCode),'GetExitCodeProcess');
    if dwExitCode <> 0 then begin
      WriteLineToLog(Format('ERROR : Command line "%s" failed!. Aborting the backup ...',[Command]));
      WriteLineToLog(Format('- Returned error code: %d',[dwExitCode]));
      raise Exception.Create(HResultToString(E_UNEXPECTED));
      end;
  finally
    // Close process and thread handles automatically when we wil leave this function
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    WriteLineToLog('-----------------------------------------------------');
    end;
  end;

// Delete all the shadow copies in the system
procedure TVolumeShadowCopy.DeleteAllSnapshots;
var
  pIEnumSnapshots : IVssEnumObject;
  hr : HResult;
  Prop : TVssObjectProp;
  ulFetched : ULong;
  lSnapshots : Long;
  idNonDeletedSnapshotID : TGuid;
begin
  // Get list all shadow copies.
  hr:=VssBackupComponents.Query(GUID_NULL,VSS_OBJECT_NONE,VSS_OBJECT_SNAPSHOT,pIEnumSnapshots);
  OleCheck(hr,'TVolumeShadowCopy.DeleteAllSnapshots:VssBackupComponents.Query');
  // If there are no shadow copies, just return
  if Succeeded(hr) then begin
    try
      // Enumerate all shadow copies. Delete each one
      while true do begin
        // Get the next element
        hr:=pIEnumSnapshots.Next(1,Prop,ulFetched);
        OleCheck(hr,'TVolumeShadowCopy.DeleteAllSnapshots:pIEnumSnapshots.Next');
        // We reached the end of list
        if ulFetched=0 then break;
        // Print the deleted shadow copy...
        with Prop.PObj.Snap do begin
          WriteLineToLog(Format('- Deleting shadow copy "%s" on %s from provider "%s" [$%.8x] ...',
              [GuidToString(m_SnapshotId),m_pwszOriginalVolumeName,
               GuidToString(m_ProviderId),m_lSnapshotAttributes]));
          // Perform the actual deletion
          hr:=VssBackupComponents.DeleteSnapshots(m_SnapshotId,VSS_OBJECT_SNAPSHOT,
              false,lSnapshots,idNonDeletedSnapshotID);
          end;
        if Failed(hr) then begin
          WriteLineToLog('Error while deleting shadow copies ...');
          WriteLineToLog(Format('- Last shadow copy that could not be deleted: "%s"',
            [GuidToString(idNonDeletedSnapshotID)]));
          OleCheck(hr,'TVolumeShadowCopy.DeleteAllSnapshots:VssBackupComponents.DeleteSnapshots');
          end;
        end;
    finally VssFreeSnapshotProperties(@Prop.PObj.Snap);
      end;
    end
  else WriteLineToLog(sLineBreak+'There are no shadow copies on the system'+sLineBreak);
  end;

// Delete the given shadow copy set
procedure TVolumeShadowCopy.DeleteSnapshotSet(snapshotSetID : TGuid);
var
  hr : HResult;
  lSnapshots : Long;
  idNonDeletedSnapshotID : TGuid;
begin
  // Print the deleted shadow copy...
  WriteLineToLog(Format('- Deleting shadow copy set "%s" ...',[GuidToString(snapshotSetID)]));
  // Perform the actual deletion
  hr:=VssBackupComponents.DeleteSnapshots(snapshotSetID,VSS_OBJECT_SNAPSHOT_SET,
        false,lSnapshots,idNonDeletedSnapshotID);
  if Failed(hr) then begin
    WriteLineToLog('Error while deleting shadow copies ...');
    WriteLineToLog(Format('- Last shadow copy that could not be deleted: "%s"',
      [GuidToString(idNonDeletedSnapshotID)]));
    OleCheck(hr,'TVolumeShadowCopy.DeleteSnapshotSet:VssBackupComponents.DeleteSnapshots');
    end;
  end;

procedure TVolumeShadowCopy.DeleteOldestSnapshot(const stringVolumeName : string);
var
  uniqueVolume : string;
  pIEnumSnapshots : IVssEnumObject;
  hr : HResult;
  Prop : TVssObjectProp;
  ulFetched : ULong;
  OldestAttributes,
  lSnapshots : Long;
  OldestSnapshotId,
  OldestProviderId,
  idNonDeletedSnapshotID : TGuid;
  OldestSnapshotTimestamp : int64;
begin
  if (not GetUniqueVolumeNameForPath(stringVolumeName,uniqueVolume)) then
    uniqueVolume:=stringVolumeName;
  // Get list all shadow copies.
  hr:=VssBackupComponents.Query(GUID_NULL,VSS_OBJECT_NONE,VSS_OBJECT_SNAPSHOT,pIEnumSnapshots);
  OleCheck(hr,'TVolumeShadowCopy.DeleteOldestSnapshot:VssBackupComponents.Query');
  // If there are no shadow copies, just return
  if Succeeded(hr) then begin
    OldestSnapshotId:=GUID_NULL;
    OldestProviderId:=GUID_NULL;
    OldestAttributes:=0;
    OldestSnapshotTimestamp:=$7FFFFFFFFFFFFFFF;
    try
      // Enumerate all shadow copies. Delete each one
      while true do begin
        // Get the next element
        hr:=pIEnumSnapshots.Next(1,Prop,ulFetched);
        OleCheck(hr,'TVolumeShadowCopy.DeleteOldestSnapshot:pIEnumSnapshots.Next');
        // We reached the end of list
        if ulFetched=0 then break;
        with Prop.PObj.Snap do if AnsiSameText(m_pwszOriginalVolumeName,uniqueVolume) and
            (OldestSnapshotTimestamp>m_tsCreationTimestamp) then begin
          OldestSnapshotId:=m_SnapshotId;
          OldestSnapshotTimestamp:=m_tsCreationTimestamp;
          OldestProviderId:=m_ProviderId;
          OldestAttributes:=m_lSnapshotAttributes;
          end;
        end;
      if OldestSnapshotId<>GUID_NULL then begin
        // Print the deleted shadow copy...
        WriteLineToLog(Format('- Deleting shadow copy "%s" on %s from provider "%s" [$%.8x] ...',
              [GuidToString(OldestSnapshotId),uniqueVolume,
               GuidToString(OldestProviderId),OldestAttributes]));
        // Perform the actual deletion
        hr:=VssBackupComponents.DeleteSnapshots(OldestSnapshotId,VSS_OBJECT_SNAPSHOT,
              false,lSnapshots,idNonDeletedSnapshotID);
        if Failed(hr) then begin
          WriteLineToLog('Error while deleting shadow copies ...');
          WriteLineToLog(Format('- Last shadow copy that could not be deleted: "%s"',
            [GuidToString(idNonDeletedSnapshotID)]));
          OleCheck(hr,'TVolumeShadowCopy.DeleteAllSnapshots:VssBackupComponents.DeleteSnapshots');
          end;
        end
      else WriteLineToLog(sLineBreak+'There are no shadow copies on the system'+sLineBreak);
    finally VssFreeSnapshotProperties(@Prop.PObj.Snap);
      end;
    end
  else WriteLineToLog(sLineBreak+'There are no shadow copies on the system'+sLineBreak);
  end;


// Delete the given shadow copy
procedure TVolumeShadowCopy.DeleteSnapshot(snapshotID : TGuid);
var
  hr : HResult;
  lSnapshots : Long;
  idNonDeletedSnapshotID : TGuid;
begin
  // Print the deleted shadow copy...
  WriteLineToLog(Format('- Deleting shadow copy "%s" ...',[GuidToString(snapshotID)]));
  // Perform the actual deletion
  hr:=VssBackupComponents.DeleteSnapshots(snapshotID,VSS_OBJECT_SNAPSHOT,
        false,lSnapshots,idNonDeletedSnapshotID);
  if Failed(hr) then begin
    WriteLineToLog('Error while deleting shadow copies ...');
    WriteLineToLog(Format('- Last shadow copy that could not be deleted: "%s"',
      [GuidToString(idNonDeletedSnapshotID)]));
    OleCheck(hr,'TVolumeShadowCopy.DeleteSnapshotSet:VssBackupComponents.DeleteSnapshots');
    end;
  end;

{ TVssThread ------------------------------------------------------------------- }
constructor TVssThread.Create (const ADrive : string; ASuspend : Boolean; APriority : TThreadPriority);
begin
  inherited Create (ASuspend);
  Priority:=APriority;
  FDrive:=ADrive; FSuccess:=false;
  FShadowDeviceName:='';
  FVolumeShadowCopy:=TVolumeShadowCopy.Create;
  end;

destructor TVssThread.Destroy;
begin
  FVolumeShadowCopy.Free;
  inherited Destroy;
  end;

function TVssThread.GetLogFileName : string;
begin
  Result:=FVolumeShadowCopy.LogFileName;
  end;

procedure TVssThread.SetLogFileName (const Value : string);
begin
  FVolumeShadowCopy.LogFileName:=Value;
  end;

function TVssThread.GetWriteLog : boolean;
begin
  Result:=FVolumeShadowCopy.WriteLog;
  end;

procedure TVssThread.SetWriteLog (Value : boolean);
begin
  FVolumeShadowCopy.WriteLog:=Value;
  end;

procedure TVssThread.SetShowStatus (ACallBack : TOutputString);
begin
  FVolumeShadowCopy.OnStatusMessage:=ACallBack;
  end;

function TVssThread.GetDone : boolean;
begin
  Result:=Terminated;
  end;

procedure TVssThread.Execute;
begin
  with FVolumeShadowCopy do begin
    try
      WriteLineToLog(Format('Starting snapshot creation for "%s"',[FDrive]));
      // Initialize the VSS client
      Initialize(VSS_CTX_BACKUP);
      // Create the shadow copy set
      CreateSnapshotSet(GetVolumeUniqueName(FDrive));
      FShadowDeviceName:=GetLastShadowDeviceName;
      FSuccess:=true;
    except
      on E:Exception do WriteLineToLog('ERROR : '+E.Message);
      end;
    end;
  Terminate;
  end;

procedure TVssThread.SaveBackupComponentsDocument(const AFileName : string);
begin
  FVolumeShadowCopy.SaveBackupComponentsDocument(AFileName);
  end;

procedure TVssThread.WriteLineToLog (const  s : string);
begin
  FVolumeShadowCopy.WriteLineToLog(s);
  end;

procedure TVssThread.DeleteShadowCopy;
begin
  FVolumeShadowCopy.BackupComplete(true);
  end;

{ Helper functions ------------------------------------------------------------------- }
// Initialize COM security
function InitSecurity : HResult;
begin
  Result:=CoInitializeSecurity(nil,
    // Allow *all* VSS writers to communicate back!
    -1,                            // Default COM authentication service
    nil,                           // Default COM authorization service
    nil,                           // reserved parameter
    RPC_C_AUTHN_LEVEL_PKT_PRIVACY, // Strongest COM authentication level
    RPC_C_IMP_LEVEL_IDENTIFY,      // Minimal impersonation abilities
    nil,                           // Default COM authentication settings
    EOAC_NONE,                     // No special options
    nil                            // Reserved parameter
    );
  end;

// Create VSS thread
function CreateVssThread (const ADrive : string; IgnoreTooLate : boolean) : TVssThread;
var
  hr : HResult;
begin
  hr:=InitSecurity;
  if succeeded(hr) or (IgnoreTooLate and (hr=RPC_E_TOO_LATE)) then
    Result:=TVssThread.Create(ADrive)
  else begin
    OleErrorHint(hr,Format('ERROR : COM call "%s" failed.',['CoInitializeSecurity']));
    Result:=nil;
    end;
  end;

end.
