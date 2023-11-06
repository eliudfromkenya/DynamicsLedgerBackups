// See https://aka.ms/new-console-template for more information
using DynamicsLedgerBackups;
using ICSharpCode.SharpZipLib.GZip;
using ICSharpCode.SharpZipLib.Zip;

Console.WriteLine("Loading up the browser to start backups");

try
{
    await Backups.CreateBackup();
    Console.WriteLine("Successfully done");
}
catch (Exception ex)
{
    Console.WriteLine(ex);
}

try
{
    var folder = SaveDownloads.folder;
    var files = Directory.GetFiles(folder, "*.xlsx");

    if (files.Any())
    {
        var date = DateTime.Now;
        var zip = Path.Combine(folder, $"backup_{date:yyyy_MM_dd_HH_mm}.zip");
        ZipMaker.Create(zip, files);
    }
}
catch (Exception)
{

	throw;
}