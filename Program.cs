using System.Diagnostics;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.CommandLine;
using FileCollector.VSS;

internal partial class Program
{

  private static EnumerationOptions DefaultEnumerationOptions() => new()
  {
    IgnoreInaccessible = true,
    AttributesToSkip = FileAttributes.ReparsePoint,
    RecurseSubdirectories = true,
    ReturnSpecialDirectories = false
  };

  private static IEnumerable<string> GetDirFiles(string dir) =>
      Directory.EnumerateFiles(dir, "*", DefaultEnumerationOptions());

  private static void AddFileToArchive(ZipArchive archive, string fpath, string dispFath)
  {
    try
    {
      archive.CreateEntryFromFile(fpath, dispFath, CompressionLevel.Fastest);
    }
    catch (Exception e)
    {
      Console.Error.WriteLine($"Error adding {dispFath} to archive: {e.Message}");
    }
  }

  private static string SanitizeQualifier(string qualifier)
  {
    return $"{qualifier[..1].ToLower()}_";
  }

  [GeneratedRegex(@"^[a-zA-Z]:\\$")]
  private static partial Regex DriveLetterRegex();

  private static (string, string, string) SanitizePath(string path)
  {
    var root = Path.GetPathRoot(path);
    if (root is not null)
    {
      if (DriveLetterRegex().IsMatch(root))
      {
        var qualif = SanitizeQualifier(root);
        var newpath = Path.Combine(qualif, path[root.Length..]);
        return (newpath, root, qualif);
      }
      else
        throw new ArgumentException($"Path {path} does not start with a drive letter.");
    }
    else
      throw new ArgumentException($"Path {path} does not have a root.");
  }

  static void CollectFiles(IEnumerable<string> toarchive, string ofilename)
  {
    var currentDir = Directory.GetCurrentDirectory();
    using var backup = new VssBackupComponents();

    var toarchiveClean = new List<string>();
    var qualifs = new Dictionary<string, string>();

    foreach (var path in toarchive)
    {
      var (spath, oldqualif, qualif) = SanitizePath(path);
      qualifs.TryAdd(oldqualif, qualif);
      toarchiveClean.Add(spath);
    }

    var qualifs_list = qualifs.Select(e => ValueTuple.Create(e.Key, e.Value));
    var sset = Guid.Empty;
    IEnumerable<Guid> sids = new List<Guid>();

    if (qualifs_list.Any())
    {
      try
      {

        (sset, sids) = backup.CreateShadowCopiesSimple(qualifs_list.Select(e => e.Item1));
      }
      catch (Exception e)
      {
        throw new Exception("Failed to create shadow copies", e);
      }
    }

    var mountedDirs = new List<string>();

    try
    {
      Debug.Assert(qualifs_list.Count() == sids.Count());

      var qualifs_snaps = qualifs_list.Zip(sids);

      foreach (var ((dletter, qualif), sid) in qualifs_snaps)
      {
        string mountPath = Path.Combine(currentDir, qualif);
        try
        {
          backup.MountSnapshot(sid, mountPath);
          mountedDirs.Add(mountPath);
        }
        catch (Exception e)
        {
          Console.Error.WriteLine($"Error: Failed to mount snapshot {sid} of drive {dletter} at {mountPath}, skipping (Cause: {e})");
        }
      }

      using var fstream = new FileStream(ofilename, FileMode.Create);
      using var archive = new ZipArchive(fstream, ZipArchiveMode.Create);
      foreach (var path in toarchiveClean)
      {
        var diskPath = Path.Combine(currentDir, path);
        if (File.Exists(diskPath))
          AddFileToArchive(archive, diskPath, path);
        else if (Directory.Exists(diskPath))
        {
          foreach (var fpath in GetDirFiles(diskPath))
            AddFileToArchive(archive, fpath, Path.GetRelativePath(currentDir, fpath));
        }
        else
          Console.Error.WriteLine($"Error: \"{path}\" is neither a valid file nor a valid directory.");
      }
    }
    finally
    {
      if (sset != Guid.Empty)
        backup.DeleteSnapshotSetSimple(sset);
      foreach (var dir in mountedDirs)
        Directory.Delete(dir);
    }
  }

  private static int Main(string[] args)
  {
    var ofilename = new Option<string>(name: "--output", description: "Output filename for the ZIP file", getDefaultValue: () => "out.zip");
    var arg = new Argument<IEnumerable<string>>(name: "paths", description: "List of files or directories to collect from the machine");
    var rootCmd = new RootCommand("File collection utility with Volume Shadow Copy Service support");
    rootCmd.AddArgument(arg);
    rootCmd.AddOption(ofilename);
    rootCmd.SetHandler((arg, ofilename) => CollectFiles(arg, ofilename), arg, ofilename);
    return rootCmd.Invoke(args);
  }
}