using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security;

namespace FileCollector.VSS;

[SuppressUnmanagedCodeSecurity()]
internal static partial class NativeMethods
{
  [LibraryImport("vssapi.dll", EntryPoint = "VssFreeSnapshotPropertiesInternal")]
  internal static partial void VssFreeSnapshotProperties(IntPtr pProp);

  [LibraryImport("vssapi.dll", EntryPoint = "CreateVssBackupComponentsInternal")]
  internal static partial int CreateVssBackupComponents(out IntPtr ptr);

  [LibraryImport("Kernel32.dll", EntryPoint = "CreateSymbolicLinkW", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  internal static partial bool CreateSymbolicLinkW(string lpSymlinkFileName, string lpTargetFileName, uint flags);
}

public enum HResult : uint
{
  SOk = 0x00000000,
  EAccessdenied = 0x80070005,
  EOutOfMemory = 0x8007000E,
  EInvalidArg = 0x80070057,
  EBadState = 0x80042301,
  EUnexpected = 0x80042302,
  EProviderAlreadyRegistered = 0x80042303,
  EProviderNotRegistered = 0x80042304,
  EProviderVeto = 0x80042306,
  EProviderInUse = 0x80042307,
  EObjectNotFound = 0x80042308,
  SAsyncPending = 0x00042309,
  SAsyncFinished = 0x0004230a,
  SAsyncCancelled = 0x0004230b,
  EVolumeNotSupported = 0x8004230c,
  EVolumeNotSupportedByProvider = 0x8004230e,
  EObjectAlreadyExists = 0x8004230d,
  EUnexpectedProviderError = 0x8004230f,
  ECorruptXmlDocument = 0x80042310,
  EInvalidXmlDocument = 0x80042311,
  EMaximumNumberOfVolumesReached = 0x80042312,
  EFlushWritesTimeout = 0x80042313,
  EHoldWritesTimeout = 0x80042314,
  EUnexpectedWriterError = 0x80042315,
  ESnapshotSetInProgress = 0x80042316,
  EMaximumNumberOfSnapshotsReached = 0x80042317,
  EWriterInfrastructure = 0x80042318,
  EWriterNotResponding = 0x80042319,
  EWriterAlreadySubscribed = 0x8004231a,
  EUnsupportedContext = 0x8004231b,
  EVolumeInUse = 0x8004231d,
  EMaximumDiffareaAssociationsReached = 0x8004231e,
  EInsufficientStorage = 0x8004231f,
  ENoSnapshotsImported = 0x80042320,
  SSomeSnapshotsNotImported = 0x00042321,
  ESomeSnapshotsNotImported = 0x80042321,
  EMaximumNumberOfRemoteMachinesReached = 0x80042322,
  ERemoteServerUnavailable = 0x80042323,
  ERemoteServerUnsupported = 0x80042324,
  ERevertInProgress = 0x80042325,
  ERevertVolumeLost = 0x80042326,
  ERebootRequired = 0x80042327,
  ETransactionFreezeTimeout = 0x80042328,
  ETransactionThawTimeout = 0x80042329,
  EVolumeNotLocal = 0x8004232d,
  EClusterTimeout = 0x8004232e,
  EWritererrorInconsistentsnapshot = 0x800423f0,
  EWritererrorOutofresources = 0x800423f1,
  EWritererrorTimeout = 0x800423f2,
  EWritererrorRetryable = 0x800423f3,
  EWritererrorNonretryable = 0x800423f4,
  EWritererrorRecoveryFailed = 0x800423f5,
  EBreakRevertIdFailed = 0x800423f6,
  ELegacyProvider = 0x800423f7,
  EMissingDisk = 0x800423f8,
  EMissingHiddenVolume = 0x800423f9,
  EMissingVolume = 0x800423fa,
  EAutorecoveryFailed = 0x800423fb,
  EDynamicDiskError = 0x800423fc,
  ENontransportableBcd = 0x800423fd,
  ECannotRevertDiskid = 0x800423fe,
  EResyncInProgress = 0x800423ff,
  EClusterError = 0x80042400,
  EUnselectedVolume = 0x8004232a,
  ESnapshotNotInSet = 0x8004232b,
  ENestedVolumeLimit = 0x8004232c,
  ENotSupported = 0x8004232f,
  EWritererrorPartialFailure = 0x80042336,
  EAsrerrorDiskAssignmentFailed = 0x80042401,
  EAsrerrorDiskRecreationFailed = 0x80042402,
  EAsrerrorNoArcpath = 0x80042403,
  EAsrerrorMissingDyndisk = 0x80042404,
  EAsrerrorSharedCridisk = 0x80042405,
  EAsrerrorDatadiskRdisk0 = 0x80042406,
  EAsrerrorRdisk0Toosmall = 0x80042407,
  EAsrerrorCriticalDisksTooSmall = 0x80042408,
  EWriterStatusNotAvailable = 0x80042409,
  EAsrerrorDynamicVhdNotSupported = 0x8004240a,
  ECriticalVolumeOnInvalidDisk = 0x80042411,
  EAsrerrorRdiskForSystemDiskNotFound = 0x80042412,
  EAsrerrorNoPhysicalDiskAvailable = 0x80042413,
  EAsrerrorFixedPhysicalDiskAvailableAfterDiskExclusion = 0x80042414,
  EAsrerrorCriticalDiskCannotBeExcluded = 0x80042415,
  EAsrerrorSystemPartitionHidden = 0x80042416,
  EFssTimeout = 0x80042417
}

public enum VssVolumeSnapshotAttributes : uint
{
  Persistent = 0x1,
  NoAutorecovery = 0x2,
  ClientAccessible = 0x4,
  NoAutoRelease = 0x8,
  NoWriters = 0x10,
  Transportable = 0x20,
  NotSurfaced = 0x40,
  NotTransacted = 0x80,
  HardwareAssisted = 0x10000,
  Differential = 0x20000,
  Plex = 0x40000,
  Imported = 0x80000,
  ExposedLocally = 0x100000,
  ExposedRemotely = 0x200000,
  Autorecover = 0x400000,
  RollbackRecovery = 0x800000,
  DelayedPostsnapshot = 0x1000000,
  TxfRecovery = 0x2000000,
  FileShare = 0x4000000
}

public enum VssSnapshotContext : uint
{
  Backup = 0,
  FileShareBackup = VssVolumeSnapshotAttributes.NoWriters,
  NasRollback = VssVolumeSnapshotAttributes.Persistent
                         | VssVolumeSnapshotAttributes.NoAutoRelease
                         | VssVolumeSnapshotAttributes.NoWriters,
  AppRollback = VssVolumeSnapshotAttributes.Persistent
                         | VssVolumeSnapshotAttributes.NoAutoRelease,
  ClientAccessible = VssVolumeSnapshotAttributes.Persistent
                              | VssVolumeSnapshotAttributes.ClientAccessible
                              | VssVolumeSnapshotAttributes.NoAutoRelease
                              | VssVolumeSnapshotAttributes.NoWriters,
  ClientAccessibleWriters = VssVolumeSnapshotAttributes.Persistent
                                      | VssVolumeSnapshotAttributes.ClientAccessible
                                      | VssVolumeSnapshotAttributes.NoAutoRelease,
  ALL = 0xffffffff
}

public enum VSS_BACKUP_TYPE : uint
{
  UNDEFINED = 0,
  FULL,
  INCREMENTAL,
  DIFFERENTIAL,
  LOG,
  COPY,
  OTHER
}

public enum VssObjectType : uint
{
  Unknown = 0,
  None,
  SnapshotSet,
  Snapshot,
  Provider,
  Count
}

public enum VssSnapshotState : uint
{
  Unknown = 0,
  Preparing,
  ProcessingPrepare,
  Prepared,
  ProcessingPrecommit,
  Precommitted,
  ProcessingCommit,
  Committed,
  ProcessingPostcommit,
  ProcessingPrefinalcommit,
  Prefinalcommitted,
  ProcessingPostfinalcommit,
  Created,
  Aborted,
  Deleted,
  Postcommitted,
  Count
}

public enum AsyncStatus
{
  Cancelled,
  Finished,
  Pending
}

static class AsyncStatusHelper
{
  public static AsyncStatus GetAsyncStatus(HResult hr) => hr switch
  {
    HResult.SAsyncCancelled => AsyncStatus.Cancelled,
    HResult.SAsyncFinished => AsyncStatus.Finished,
    HResult.SAsyncPending => AsyncStatus.Pending,
    _ => throw new Exception($"Could not get async status: {hr}"),
  };
}

public class VssAsync : IDisposable
{
  // COM IID
  private static readonly Guid IID_IVssAsync = new("507C37B4-CF5B-4e95-B0AF-14EB9767467E");

  // VTable offsets
  private const uint OFF_Cancel = 3;
  private const uint OFF_Wait = 4;
  private const uint OFF_QueryStatus = 5;

  private readonly IntPtr ivssAsyncInst;
  private bool disposedValue = false;

  // Taking ownership of ivssAsyncInst
  public VssAsync(IntPtr ivssAsyncInst)
  {
    try
    {
      var idcp = IID_IVssAsync;
      int hr = Marshal.QueryInterface(ivssAsyncInst, ref idcp, out this.ivssAsyncInst);
      if (hr != 0)
        throw new Exception($"Query interface failed: {(HResult)hr}");
    }
    finally
    {
      Marshal.Release(ivssAsyncInst);
    }
  }

  ~VssAsync()
  {
    Dispose();
  }

  public void Cancel()
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, int>)(*(*(void***)ivssAsyncInst + OFF_Cancel)))(ivssAsyncInst);
      if (hr != 0)
        throw new Exception($"Failed to cancel IVssAsync: {(HResult)hr}");
    }
  }

  private void DisposeInternal()
  {
    if (!disposedValue)
    {
      Marshal.Release(ivssAsyncInst);
      disposedValue = true;
    }
  }

  public void Dispose()
  {
    DisposeInternal();
    GC.SuppressFinalize(this);
  }

  public AsyncStatus QueryStatus()
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, out uint, out int, int>)(*(*(void***)ivssAsyncInst + OFF_QueryStatus)))(ivssAsyncInst, out uint pHrResult, out int _);
      if (hr != 0)
        throw new Exception($"Failed to query IVssAsync status: {(HResult)hr}");
      return AsyncStatusHelper.GetAsyncStatus((HResult)pHrResult);
    }
  }

  public void Wait()
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, int>)(*(*(void***)ivssAsyncInst + OFF_Wait)))(ivssAsyncInst);
      if (hr != 0)
        throw new Exception($"Failed to query IVssAsync status: {(HResult)hr}");
    }
  }

  public void RunToCompletion()
  {
    try
    {
      Wait();
      Debug.Assert(QueryStatus() == AsyncStatus.Finished);
    }
    finally
    {
      Dispose();
    }
  }
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct VssSnapshotProperties
{
  public Guid m_SnapshotId;
  public Guid m_SnapshotSetId;
  public uint m_lSnapshotsCount;
  public string m_pwszSnapshotDeviceObject;
  public string m_pwszOriginalVolumeName;
  public string m_pwszOriginatingMachine;
  public string m_pwszServiceMachine;
  public string m_pwszExposedName;
  public string m_pwszExposedPath;
  public Guid m_ProviderId;
  public uint m_lSnapshotAttributes;
  public ulong m_tsCreationTimestamp;
  public VssSnapshotState m_eStatus;

  public override readonly string ToString()
  {
    return @$"m_SnapshotId: {m_SnapshotId}, 
m_SnapshotSetId: {m_SnapshotSetId},
m_lSnapshotsCount: {m_lSnapshotsCount},
m_pwszSnapshotDeviceObject: {m_pwszSnapshotDeviceObject},
m_pwszOriginalVolumeName: {m_pwszOriginalVolumeName},
m_pwszOriginatingMachine: {m_pwszOriginatingMachine},
m_pwszServiceMachine: {m_pwszServiceMachine},
m_pwszServiceMachine: {m_pwszServiceMachine},
m_pwszExposedName: {m_pwszExposedName},
m_pwszExposedPath: {m_pwszExposedPath},
m_ProviderId: {m_ProviderId},
m_lSnapshotAttributes: {m_lSnapshotAttributes},
m_tsCreationTimestamp: {m_tsCreationTimestamp},
m_eStatus: {m_eStatus}";
  }
}

public record SnapShotDeleteInfo(uint NumDeteletedShadowCopies, Guid NonDeletedSnapshotID);

public class VssBackupComponents : IDisposable
{
  // COM IID
  private static readonly Guid IID_IVssBackupComponents = new("665c1d5f-c218-414d-a05d-7fef5f9d5c86");

  // VTable offsets
  private const uint OFF_GetWriterComponentsCount = 3;
  private const uint OFF_GetWriterComponents = 4;
  private const uint OFF_InitializeForBackup = 5;
  private const uint OFF_SetBackupState = 6;
  private const uint OFF_InitializeForRestore = 7;
  private const uint OFF_SetRestoreState = 8;
  private const uint OFF_GatherWriterMetadata = 9;
  private const uint OFF_GetWriterMetadataCount = 10;
  private const uint OFF_GetWriterMetadata = 11;
  private const uint OFF_FreeWriterMetadata = 12;
  private const uint OFF_AddComponent = 13;
  private const uint OFF_PrepareForBackup = 14;
  private const uint OFF_AbortBackup = 15;
  private const uint OFF_GatherWriterStatus = 16;
  private const uint OFF_GetWriterStatusCount = 17;
  private const uint OFF_FreeWriterStatus = 18;
  private const uint OFF_GetWriterStatus = 19;
  private const uint OFF_SetBackupSucceeded = 20;
  private const uint OFF_SetBackupOptions = 21;
  private const uint OFF_SetSelectedForRestore = 22;
  private const uint OFF_SetRestoreOptions = 23;
  private const uint OFF_SetAdditionalRestores = 24;
  private const uint OFF_SetPreviousBackupStamp = 25;
  private const uint OFF_SaveAsXML = 26;
  private const uint OFF_BackupComplete = 27;
  private const uint OFF_AddAlternativeLocationMapping = 28;
  private const uint OFF_AddRestoreSubcomponent = 29;
  private const uint OFF_SetFileRestoreStatus = 30;
  private const uint OFF_AddNewTarget = 31;
  private const uint OFF_SetRangesFilePath = 32;
  private const uint OFF_PreRestore = 33;
  private const uint OFF_PostRestore = 34;
  private const uint OFF_SetContext = 35;
  private const uint OFF_StartSnapshotSet = 36;
  private const uint OFF_AddToSnapshotSet = 37;
  private const uint OFF_DoSnapshotSet = 38;
  private const uint OFF_DeleteSnapshots = 39;
  private const uint OFF_ImportSnapshots = 40;
  private const uint OFF_BreakSnapshotSet = 41;
  private const uint OFF_GetSnapshotProperties = 42;
  private const uint OFF_Query = 43;
  private const uint OFF_IsVolumeSupported = 44;
  private const uint OFF_DisableWriterClasses = 45;
  private const uint OFF_EnableWriterClasses = 46;
  private const uint OFF_DisableWriterInstances = 47;
  private const uint OFF_ExposeSnapshot = 48;
  private const uint OFF_RevertToSnapshot = 49;
  private const uint OFF_QueryRevertStatus = 50;
  private readonly IntPtr ivssBackupComponentsInst;
  private bool disposedValue = false;
  private bool dirtySnapshotSet = false;

  public VssBackupComponents()
  {
    int hr = NativeMethods.CreateVssBackupComponents(out IntPtr vssinstance);
    try
    {
      if (hr != 0)
      {
        throw new Exception($"Failed to create component: {(HResult)hr}");
      }
      var idcp = IID_IVssBackupComponents;
      hr = Marshal.QueryInterface(vssinstance, ref idcp, out IntPtr ivssBackupComponentsInst);
      if (hr != 0)
        throw new Exception($"Query interface failed: {(HResult)hr}");
      this.ivssBackupComponentsInst = ivssBackupComponentsInst;
    }
    finally
    {
      Marshal.Release(vssinstance);
    }
  }

  private unsafe VssAsync InvokeSimpleAsyncMethod(uint methodOffset)
  {
    int hr = ((delegate* unmanaged<IntPtr, out IntPtr, int>)(*(*(void***)ivssBackupComponentsInst + methodOffset)))(ivssBackupComponentsInst, out IntPtr ivssAsyncInst);
    if (hr != 0)
      throw new Exception($"Failed to create IVssAsync instance: {(HResult)hr}");
    return new(ivssAsyncInst);
  }

  static R WithBStr<R>(string? s, Func<IntPtr, R> f)
  {
    var bstr = Marshal.StringToBSTR(s);
    try
    {
      return f(bstr);
    }
    finally
    {
      Marshal.FreeBSTR(bstr);
    }
  }

  static R WithCoTaskMemString<R>(string? s, Func<IntPtr, R> f)
  {
    var str = Marshal.StringToCoTaskMemUni(s);
    try
    {
      return f(str);
    }
    finally
    {
      Marshal.FreeCoTaskMem(str);
    }
  }

  public void AbortBackup()
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, int>)(*(*(void***)ivssBackupComponentsInst + OFF_AbortBackup)))(ivssBackupComponentsInst);
      if (hr != 0)
        throw new Exception($"Failed to abort backup: {(HResult)hr}");
    }
  }

  public Guid AddToSnapshotSet(string volumeName) => WithCoTaskMemString(volumeName, pwszVolumeName =>
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, IntPtr, Guid, out Guid, int>)(*(*(void***)ivssBackupComponentsInst + OFF_AddToSnapshotSet)))(
        ivssBackupComponentsInst, pwszVolumeName, Guid.Empty, out Guid pidSnapshot);
      if (hr != 0)
        throw new Exception($"Failed to intialize backup: {(HResult)hr}");
      return pidSnapshot;
    }
  });

  public VssAsync BackupComplete() => InvokeSimpleAsyncMethod(OFF_BackupComplete);

  public SnapShotDeleteInfo DeleteSnapshots(Guid sourceObjectId, VssObjectType eSourceObjectType, bool bForceDelete)
  {
    if (eSourceObjectType != VssObjectType.Snapshot && eSourceObjectType != VssObjectType.SnapshotSet)
      throw new ArgumentException("Object passed to DeleteSnapshots must be a snapshot or a snapshot set");
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, Guid, uint, bool, out uint, out Guid, int>)(*(*(void***)ivssBackupComponentsInst + OFF_DeleteSnapshots)))(
        ivssBackupComponentsInst, sourceObjectId, (uint)eSourceObjectType, bForceDelete, out uint NumDeteletedShadowCopies, out Guid NonDeletedSnapshotID);
      if (hr != 0)
        throw new Exception($"Failed to delete snapshot or snapshotset: {(HResult)hr}");
      return new(NumDeteletedShadowCopies, NonDeletedSnapshotID);
    }
  }

  public VssAsync DoSnapshotSet()
  {
    var asyncw = InvokeSimpleAsyncMethod(OFF_DoSnapshotSet);
    dirtySnapshotSet = false;
    return asyncw;
  }

  public void FreeWriterMetadata()
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, int>)(*(*(void***)ivssBackupComponentsInst + OFF_FreeWriterMetadata)))(ivssBackupComponentsInst);
      if (hr != 0)
        throw new Exception($"Failed to free writer metadata: {(HResult)hr}");
    }
  }

  public VssAsync GatherWriterMetadata() => InvokeSimpleAsyncMethod(OFF_GatherWriterMetadata);

  public VssAsync GatherWriterStatus() => InvokeSimpleAsyncMethod(OFF_GatherWriterStatus);

  public VssSnapshotProperties GetSnapshotProperties(Guid SnapshotId)
  {
    unsafe
    {
      IntPtr pProp = Marshal.AllocCoTaskMem(Marshal.SizeOf<VssSnapshotProperties>());
      try
      {
        int hr = ((delegate* unmanaged<IntPtr, Guid, IntPtr, int>)(*(*(void***)ivssBackupComponentsInst + OFF_GetSnapshotProperties)))(
          ivssBackupComponentsInst, SnapshotId, pProp);
        if (hr != 0)
          throw new Exception($"Failed to get snapshort properties: {(HResult)hr}");
        var props = Marshal.PtrToStructure<VssSnapshotProperties>(pProp);
        NativeMethods.VssFreeSnapshotProperties(pProp);
        return props;
      }
      finally
      {
        Marshal.FreeCoTaskMem(pProp);
      }
    }
  }

  public void InitializeForBackup(string? bstrXML = null) => WithBStr<object>(bstrXML, bstrXML =>
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, IntPtr, int>)(*(*(void***)ivssBackupComponentsInst + OFF_InitializeForBackup)))(ivssBackupComponentsInst, bstrXML);
      if (hr != 0)
        throw new Exception($"Failed to intialize backup: {(HResult)hr}");
#pragma warning disable CS8603
      return null;
#pragma warning restore CS8603
    }
  });

  public bool IsVolumeSupported(string pwszVolumeName) => WithCoTaskMemString(pwszVolumeName, pwszVolumeName =>
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, Guid, IntPtr, out bool, int>)(*(*(void***)ivssBackupComponentsInst + OFF_IsVolumeSupported)))(
        ivssBackupComponentsInst, Guid.Empty, pwszVolumeName, out bool volumeSupported);
      if (hr != 0)
        throw new Exception($"Failed to check if volume is supported: {(HResult)hr}");
      return volumeSupported;
    }
  });

  public VssAsync PrepareForBackup() => InvokeSimpleAsyncMethod(OFF_PrepareForBackup);

  public void SetBackupState(bool bSelectComponents, bool bBackupBootableSystemState, VSS_BACKUP_TYPE backupType, bool bPartialFileSupport = false)
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, bool, bool, uint, bool, int>)(*(*(void***)ivssBackupComponentsInst + OFF_SetBackupState)))(
        ivssBackupComponentsInst, bSelectComponents, bBackupBootableSystemState, (uint)backupType, bPartialFileSupport);
      if (hr != 0)
        throw new Exception($"Failed to set backup state: {(HResult)hr}");
    }
  }

  public void SetContext(uint lContext)
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, uint, int>)(*(*(void***)ivssBackupComponentsInst + OFF_SetContext)))(ivssBackupComponentsInst, lContext);
      if (hr != 0)
        throw new Exception($"Failed to set context: {(HResult)hr}");
    }
  }

  public Guid StartSnapshotSet()
  {
    unsafe
    {
      int hr = ((delegate* unmanaged<IntPtr, out Guid, int>)(*(*(void***)ivssBackupComponentsInst + OFF_StartSnapshotSet)))(ivssBackupComponentsInst, out Guid pSnapshotSetId);
      if (hr != 0)
        throw new Exception($"Could not create snapshot set: {(HResult)hr}");
      dirtySnapshotSet = true;
      return pSnapshotSetId;
    }
  }

  private void DisposeInternal()
  {
    if (!disposedValue)
    {
      if (dirtySnapshotSet)
        AbortBackup();
      Marshal.Release(ivssBackupComponentsInst);
      disposedValue = true;
    }
  }

  public void MountSnapshot(Guid snapid, string mountAt)
  {
    var props = GetSnapshotProperties(snapid);
    var devObjPath = props.m_pwszSnapshotDeviceObject + @"\";
    if (!NativeMethods.CreateSymbolicLinkW(mountAt, devObjPath, 0x1))
    {
      var errcode = Marshal.GetLastPInvokeError();
      var errmsg = Marshal.GetLastPInvokeErrorMessage();
      throw new Exception($"Failed to create symbolic link from {devObjPath} to {mountAt}, code: {errcode} \"{errmsg}\"");
    }
  }

  public (Guid, IEnumerable<Guid>) CreateShadowCopiesSimple(IEnumerable<string> volumes)
  {
    var guids = new List<Guid>();
    InitializeForBackup();
    SetBackupState(false, false, VSS_BACKUP_TYPE.COPY);
    GatherWriterMetadata().RunToCompletion();
    SetContext((uint)VssSnapshotContext.ClientAccessibleWriters);
    var sset = StartSnapshotSet();
    guids.AddRange(volumes.Select(AddToSnapshotSet));
    PrepareForBackup().RunToCompletion();
    DoSnapshotSet().RunToCompletion();
    GatherWriterStatus().RunToCompletion();
    BackupComplete().RunToCompletion();
    return (sset, guids);
  }

  public void DeleteSnapshotSetSimple(Guid sset)
  {
    var (_, failuid) = DeleteSnapshots(sset, VssObjectType.SnapshotSet, true);
    Debug.Assert(failuid == Guid.Empty);
  }

  ~VssBackupComponents()
  {
    Dispose();
  }

  public void Dispose()
  {
    DisposeInternal();
    GC.SuppressFinalize(this);
  }
}
