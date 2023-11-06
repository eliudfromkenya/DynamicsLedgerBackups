using Microsoft.Playwright;
using System;
using System.Text.RegularExpressions;

namespace DynamicsLedgerBackups
{
    internal class Backups
    {
        private static IPage[] pages = new IPage[7];
        //static async Task Refresh()
        //{
        //    Task.Run(() =>
        //    {
        //        try
        //        {
        //            while (true)
        //            {
        //                page.GetByRole(AriaRole.List).ClickAsync();
        //                Thread.Sleep(10000);
        //            }
        //        }
        //        catch { }
        //    });
        //}

        public static async Task CreateBackup()
        {
            using var playwright = await Playwright.CreateAsync();
            string url = "";
            string username="", password = "";
            await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = false
            });

            var context = await browser.NewContextAsync(new()
            {
                HttpCredentials = new HttpCredentials
                {
                    Username = username,
                    Password = password
                },
                ViewportSize = null
            });
            context.SetDefaultTimeout(400000000);

            for (int i = 0; i < 7; i++)
            {
                pages[i] = await context.NewPageAsync();
                await pages[i].GotoAsync(url);
            }

            var queue = new Queue<Func<IPage, Task>>(Tasks());
            void LoadPage(IPage page)
            {
                Func<IPage, Task> task;
                lock (pages)
                {
                    if (queue.Any())
                        task = queue.Dequeue();
                    else
                    {
                        page.CloseAsync();
                        return;
                    }
                }
                page.SetDefaultTimeout(400000000);
                Task val;
                lock (pages)
                {
                    val = task(page);
                    page.BringToFrontAsync();
                }

                using (val)
                {
                    try
                    {
                        val.Wait();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                    if (val.IsCompleted)
                        LoadPage(page);
                }
            }
            var tsks = pages
                .Select(pg => Task.Run(() => LoadPage(pg)))
                .ToArray();
            Task.WaitAll(tsks);
            foreach (var tsk in tsks)
                tsk.Dispose();
        }

        public static Func<IPage, Task>[] Tasks()
        {
            List<Func<IPage, Task>> tasks = new();

            tasks.Add(page => Task.Run(async () =>
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync();

                await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("vendors");

                await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Vendors", new() { Exact = true }).First.ClickAsync();

                var download = await page.RunAndWaitForDownloadAsync(async () =>
                {
                    await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync();
                });
                await SaveDownloads.SaveFile(download);
            }));

            tasks.Add(page => Task.Run(async () =>
           {
               // await page.GetByRole(AriaRole.List).ClickAsync();

               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync();

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("customers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Customers", new() { Exact = true }).First.ClickAsync();

               var download1 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync();
               });
               await SaveDownloads.SaveFile(download1);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync();

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("items");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Items", new() { Exact = true }).First.ClickAsync();

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Collapse the FactBox pane" }).ClickAsync();

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Show the rest" }).ClickAsync();

               var download1 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync();
               });
               await SaveDownloads.SaveFile(download1);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("FA Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });

               await SaveDownloads.SaveFile(download);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Job Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download1 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });

               await SaveDownloads.SaveFile(download1);
           }));

            tasks.Add(page => Task.Run(async () =>
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

                await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

                await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

                await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Item Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 30000000 });

                var download = await page.RunAndWaitForDownloadAsync(async () =>
                {
                    await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
                });
                await SaveDownloads.SaveFile(download);
            }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Check Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download3 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download3);
           }));
            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Vendor Ledger Entries", new() { Exact = true }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download);
           }));
            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("General Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download);
           }));
            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Service Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download6 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download6);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Capacity Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download7 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download7);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Employee Ledger Entries", new() { Exact = true }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download8 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download8);
           }));
            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Resource Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download9 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download9);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Customer Ledger Entries", new() { Exact = true }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Warranty Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download11 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download11);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Cash Flow Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download12 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download12);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Maintenance Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download13 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download13);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Bank Account Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Day Book Vendor Ledger Entry").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Cancel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Detailed Vendor Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download15 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download15);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Detailed Employee Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download16 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download16);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Detailed Customer Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Insurance Coverage Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download18 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download18);
           }));
            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("ledgers");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (22)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Physical Inventory Ledger Entries").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download19 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download19);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("Posted sales invoices");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Posted Sales Invoices").Nth(2).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download20 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download20);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.FrameLocator("iframe[title=\"Main Content\"]").Locator("div").Filter(new() { HasTextRegex = new Regex("^Posted Sales Invoices$") }).First.ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("posted sales invoices");

               await page.GotoAsync("https://mis.kenyafarmersassociation.co.ke/BC130/?bookmark=25%3bcAAAAAJ7%2f0MAUwAtADEAMAAxADQ%3d&page=143&company=KFA&dc=0");

               var download = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("posted purchase invoices");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Posted Purchase Invoices", new() { Exact = true }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               var download22 = await page.RunAndWaitForDownloadAsync(async () =>
               {
                   await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Menuitem, new() { Name = "Open in Excel" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
               });
               await SaveDownloads.SaveFile(download22);
           }));

            tasks.Add(page => Task.Run(async () =>
           {
               await page.GetByRole(AriaRole.Button, new() { Name = "" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByLabel("Type to start search:").FillAsync("posted receipts");

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByRole(AriaRole.Button, new() { Name = "Go to Reports and Analysis Show all (25)" }).ClickAsync(new LocatorClickOptions { Timeout = 3000000 });

               await page.FrameLocator("iframe[title=\"Main Content\"]").GetByText("Posted Return Receipt").ClickAsync(new LocatorClickOptions { Timeout = 3000000 });
           }));
            return tasks.ToArray();
        }
    }
}