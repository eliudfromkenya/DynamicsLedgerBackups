using Microsoft.Playwright;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicsLedgerBackups
{
    internal static class SaveDownloads
    {
        public static string folder = "Backups";
        public async static Task SaveFile(IDownload download)
        {
            // Wait for the download process to complete and save the downloaded file somewhere
            await download.SaveAsAsync(Path.Combine(folder, download.SuggestedFilename));
        }

        static SaveDownloads()
        {
            folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Dynamics Backups" );
            if(!Directory.Exists(folder)) 
            {  Directory.CreateDirectory(folder); }
        }
    }
}
