/****************************** Module Header ******************************\
Module Name:  Connect.cs
Project:      CSOneNoteRibbonAddIn
Copyright (c) Microsoft Corporation.

Hosts the event notifications that occur to add-ins, such as when they are loaded, 
unloaded, updated, and so forth.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using CSOneNoteRibbonAddIn;
using System;
using System.Windows.Forms;

namespace CSOneNoteRibbonAddIn
{
    #region Imports directives
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.OneNote;
    using System;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text.Json;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using static CSOneNoteRibbonAddIn.BookMark_Window;
    using OneNote = Microsoft.Office.Interop.OneNote;

    #endregion

    #region Read me for Add-in installation and setup information.
    // When run, the Add-in wizard prepared the registry for the Add-in.
    // At a later time, if the Add-in becomes unavailable for reasons such as:
    //   1) You moved this project to a computer other than which is was originally created on.
    //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
    //   3) Registry corruption.
    // you will need to re-register the Add-in by building the CSOneNoteRibbonAddInSetup project, 
    // right click the project in the Solution Explorer, then choose install.
    #endregion


    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [GuidAttribute("0BE84534-48A5-48A7-A9BD-0B5CAE7E12A0"),
    ProgId("CSOneNoteRibbonAddIn.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2, IRibbonExtensibility
    {
        private object applicationObject;
        private object addInInstance;
        private IRibbonUI ribbon;
        private Thread _uiThreadBookmark;
        private Thread _uiThreadNotes;
        private BookMark_Window _bookmarkWindow;
        private Option_Window _notesWindow;
        private CancellationTokenSource cts = new CancellationTokenSource();

        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public Connect()
        {
        }

        /// <summary>
        ///      Implements the OnConnection method of the IDTExtensibility2 interface.
        ///      Receives notification that the Add-in is being loaded.
        /// </summary>
        /// <param term='application'>
        ///      Root object of the host application.
        /// </param>
        /// <param term='connectMode'>
        ///      Describes how the Add-in is being loaded.
        /// </param>
        /// <param term='addInInst'>
        ///      Object representing this Add-in.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode,
            object addInInst, ref System.Array custom)
        {
            //MessageBox.Show("CSOneNoteRibbonAddIn OnConnection UPDATE");
            applicationObject = application;
            addInInstance = addInInst;
        }

        /// <summary>
        ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
        ///     Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param term='disconnectMode'>
        ///      Describes how the Add-in is being unloaded.
        /// </param>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, 
            ref System.Array custom)
        {
            //MessageBox.Show("CSOneNoteRibbonAddIn OnDisconnection");
            try
            {
                _bookmarkWindow.Activate();
                while (System.Windows.Forms.Application.OpenForms.Count > 0)
                {
                    System.Windows.Forms.Application.OpenForms[0].Close();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("closing app");
            }
            _uiThreadBookmark.Join(5000);
            _uiThreadBookmark = null;
            this.applicationObject = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref System.Array custom)
        {
            //MessageBox.Show("CSOneNoteRibbonAddIn OnAddInsUpdate");
        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
            Task backupTask = AutoExportHelper.RunPeriodicCopyAsync(cts.Token);
            /// Run the form on the UI thread
            //MessageBox.Show("CSOneNoteRibbonAddIn OnStartupComplete");
        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref System.Array custom)
        {
            //MessageBox.Show("CSOneNoteRibbonAddIn OnBeginShutdown");
            cts.Cancel();
            if (this.applicationObject != null)
            {
                this.applicationObject = null;
            }
        }

        /// <summary>
        ///     Loads the XML markup from an XML customization file 
        ///     that customizes the Ribbon user interface.
        /// </summary>
        /// <param name="RibbonID">The ID for the RibbonX UI</param>
        /// <returns>string</returns>
        /// 


        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
                  <ribbon>
                    <tabs>
                      <tab idMso='TabHome'>
                        <group id='customGroup' label='BookMark'>
                          <button id='showFormButton'
                                  label='Saved'
                                  imageMso='BookmarkInsert'
                                  size='large'
                                  onAction='OnShowFormButtonClick' />
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        #region Button Handler
        public BookMark_Window CreateBookmarkWindow(
         string selectedId,
         string selectedScope,
         string displayText,
         string notebookName,
         string notebookColor,
         string sectionGroupName,
         string sectionName,
         string sectionColor,
         string pageName,
         string paraContent)
        {
            return new BookMark_Window(
                selectedId,
                selectedScope,
                displayText,
                notebookName,
                notebookColor,
                sectionGroupName,
                sectionName,
                sectionColor,
                pageName,
                paraContent)
            {
                StartPosition = FormStartPosition.Manual,
                TopMost = true
            };
        }
        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }
        public void OnShowFormButtonClick(IRibbonControl control)
        {
            try
            {
                var model = new AddInModel();

                // Extract data from your model
                string selectedId = model.Page?.Id ?? "";
                string selectedScope = "page";
                string displayText = model.Page?.Name ?? "";
                string notebookName = model.NotebookName ?? "";
                string notebookColor = model.NotebookColor ?? "";
                string sectionGroupName = model.SectionGroup?.Name ?? "";
                string sectionName = model.Section?.Name ?? "";
                string sectionColor = model.Section?.Color ?? "";
                string pageName = model.Page?.Name ?? "";
                string paraContent = model.Page?.Paragraphs?.FirstOrDefault()?.Name ?? "";

                //----BOOKMARK WINDOW----
                if (_uiThreadBookmark != null && _uiThreadBookmark.IsAlive && _bookmarkWindow != null)
                {
                    _bookmarkWindow.Invoke((Action)(() =>
                    {
                        _bookmarkWindow.UpdateBookmarkInfo(
                            selectedId, selectedScope, displayText,
                            notebookName, notebookColor,
                            sectionGroupName, sectionName, sectionColor,
                            pageName, paraContent);
                        PositionFormNearCursor(_bookmarkWindow);
                        _bookmarkWindow.Show();
                        _bookmarkWindow.Activate();
                        _bookmarkWindow.BringToFront();
                    }));
                }
                else
                {
                    _uiThreadBookmark = new Thread(() =>
                    {
                        try
                        {
                            InitializeWindowsForms();
                            _bookmarkWindow = CreateBookmarkWindow(
                                selectedId, selectedScope, displayText,
                                notebookName, notebookColor,
                                sectionGroupName, sectionName, sectionColor,
                                pageName, paraContent);

                            PositionFormNearCursor(_bookmarkWindow);
                            System.Windows.Forms.Application.Run(_bookmarkWindow);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error launching bookmark window: " + ex.Message);
                        }
                    });
                    _uiThreadBookmark.SetApartmentState(ApartmentState.STA);
                    _uiThreadBookmark.Start();
                }

                //// ---- SECOND WINDOW (Options Window) ----
                //if (_uiThreadNotes != null && _uiThreadNotes.IsAlive && _notesWindow != null)
                //{
                //    _notesWindow.Invoke((Action)(() =>
                //    {
                //        _notesWindow.Show();
                //        _notesWindow.Activate();
                //        _notesWindow.BringToFront();
                //    }));
                //}
                //else
                //{
                //    _uiThreadNotes = new Thread(() =>
                //    {
                //        try
                //        {
                //            InitializeWindowsForms();
                //            _notesWindow = CreateNotesWindow(_bookmarkWindow);
                //            PositionFormNearCursor(_notesWindow, offsetX: 0);

                //            System.Windows.Forms.Application.Run(_notesWindow);
                //        }
                //        catch (Exception ex)
                //        {
                //            MessageBox.Show("Error launching notes window: " + ex.Message);
                //        }
                //    });
                //    _uiThreadNotes.SetApartmentState(ApartmentState.STA);
                //    _uiThreadNotes.Start();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected error: " + ex.Message);
            }
        }
        #endregion

        #region Helper Methods
        public AddInModel GetCurrentNotebookModel(OneNote.Application oneNoteApp)
        {
            var model = new AddInModel();

            try
            {
                // Get the entire hierarchy with pages
                string hierarchyXml;
                oneNoteApp.GetHierarchy("", OneNote.HierarchyScope.hsPages, out hierarchyXml);

                var doc = new System.Xml.XmlDocument();
                doc.LoadXml(hierarchyXml);

                var nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                // Get current page id
                var window = oneNoteApp.Windows.CurrentWindow;
                string currentPageId = window.CurrentPageId;

                // Find current page node
                var pageNode = doc.SelectSingleNode($"//one:Page[@ID='{currentPageId}']", nsmgr);
                if (pageNode == null)
                    return null; // current page not found

                // Get current section node (parent of page)
                var sectionNode = pageNode.ParentNode;
                if (sectionNode == null || sectionNode.Name != "one:Section")
                    return null;

                // Get current notebook node (ancestor notebook)
                var notebookNode = sectionNode.ParentNode;

                // If section is inside SectionGroup, parent is SectionGroup, notebook is one level above
                SectionGroupModel sectionGroupModel = null;
                if (notebookNode.Name == "one:SectionGroup")
                {
                    var sectionGroupNode = notebookNode;
                    notebookNode = sectionGroupNode.ParentNode;

                    sectionGroupModel = new SectionGroupModel
                    {
                        Id = sectionGroupNode.Attributes["ID"]?.Value,
                        Name = sectionGroupNode.Attributes["name"]?.Value,
                    };

                    model.SectionGroup = sectionGroupModel;
                }

                if (notebookNode == null || notebookNode.Name != "one:Notebook")
                    return null;

                // Fill notebook info
                model.NotebookId = notebookNode.Attributes["ID"]?.Value;
                model.NotebookName = notebookNode.Attributes["name"]?.Value;
                model.NotebookColor = notebookNode.Attributes["color"]?.Value;

                // Fill current section info
                var sectionModel = new SectionModel
                {
                    Id = sectionNode.Attributes["ID"]?.Value,
                    Name = sectionNode.Attributes["name"]?.Value,
                    Color = sectionNode.Attributes["color"]?.Value
                };
                model.Section = sectionModel;

                // Fill current page info
                var pageModel = new PageModel
                {
                    Id = pageNode.Attributes["ID"]?.Value,
                    Name = pageNode.Attributes["name"]?.Value
                };

                LoadParagraphs(oneNoteApp, pageModel); // load paragraphs into current page
                model.Page = pageModel;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading OneNote hierarchy: {ex.Message}", "Error", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                throw;
            }

            return model;
        }
        private void LoadParagraphs(OneNote.Application oneNoteApp, PageModel page)
        {
            using (MethodTimerLog.Time("LoadParagraphs"))
            {
                try
                {
                    string pageXml;
                    oneNoteApp.GetPageContent(page.Id, out pageXml, PageInfo.piAll);

                    var doc = new System.Xml.XmlDocument();
                    doc.LoadXml(pageXml);

                    var nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                    nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                    // First try: get the paragraph that was last selected
                    var selectedParaNode = doc.SelectSingleNode("//one:OE[@selected]/one:T", nsmgr);

                    // Fallback: if nothing selected, take the first paragraph
                    if (selectedParaNode == null)
                    {
                        selectedParaNode = doc.SelectSingleNode("//one:Outline/one:OEChildren/one:OE/one:T", nsmgr);
                    }

                    if (selectedParaNode != null)
                    {
                        string paraText = selectedParaNode.InnerText?.Trim();
                        if (!string.IsNullOrEmpty(paraText))
                        {
                            page.Paragraphs.Add(new ParagraphModel
                            {
                                Id = page.Id + "_selected",
                                Name = paraText
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading paragraphs for page {page.Name}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw;
                }
            }
        }
        #endregion
        public void InitializeWindowsForms()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
        }
        public void PositionFormNearCursor(Form form)
        {
            const string sizeFilePath = "form_size.txt";
            const int defaultWidth = 500;
            const int defaultHeight = 400;

            // Load size from file or use default
            int formWidth = defaultWidth;
            int formHeight = defaultHeight;

            try
            {
                if (File.Exists(sizeFilePath))
                {
                    string sizeText = File.ReadAllText(sizeFilePath);
                    string[] parts = sizeText.Split(',');

                    if (parts.Length == 2 &&
                        int.TryParse(parts[0].Trim(), out int width) &&
                        int.TryParse(parts[1].Trim(), out int height) &&
                        width > 0 && height > 0)
                    {
                        formWidth = width;
                        formHeight = height;
                    }
                }
            }
            catch
            {
                // Ignore errors, use defaults
            }

            // Position form directly to the right of the cursor, no boundary checks
            var cursorPos = Cursor.Position;
            int x = cursorPos.X + 1; // +1 pixel padding
            int y = cursorPos.Y;

            form.Size = new Size(formWidth, formHeight);
            form.Location = new Point(x, y);

            // Update size file on size change (avoid multiple subscriptions)
            form.SizeChanged -= (s, e) => { }; // dummy remove all
            form.SizeChanged += (s, e) =>
            {
                try
                {
                    File.WriteAllText(sizeFilePath, $"{form.Width},{form.Height}");
                }
                catch
                {
                    // Ignore write errors
                }
            };

            form.Show();
        }

        #region Soon to be deprecated methods

        public void GetSelectedOutlineInfo(OneNote.Application oneNoteApp, out string selectedId, out string selectedScope)
        {
            selectedId = null;
            selectedScope = "Paragraph";

            try
            {
                string hierarchyXml;
                oneNoteApp.GetHierarchy("", HierarchyScope.hsPages, out hierarchyXml);

                var window = oneNoteApp.Windows.CurrentWindow;
                string currentPageId = window.CurrentPageId;

                string pageXml;
                oneNoteApp.GetPageContent(currentPageId, out pageXml, PageInfo.piAll);

                var doc = new System.Xml.XmlDocument();
                doc.LoadXml(pageXml);

                var nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                var selectedOutline = doc.SelectSingleNode("//one:Outline[@selected='true']", nsmgr);
                if (selectedOutline != null)
                {
                    selectedId = selectedOutline.Attributes["ID"]?.Value;
                }
                else
                {
                    var firstOutline = doc.SelectSingleNode("//one:Outline", nsmgr);
                    selectedId = firstOutline?.Attributes["ID"]?.Value;
                }

                if (string.IsNullOrEmpty(selectedId))
                {
                    selectedId = currentPageId;
                    selectedScope = "Page";
                }
            }
            catch
            {
                // Fallback to page-level if problem occurs
                selectedId = null;
                selectedScope = "Page";
            }
        }
        /// <summary>
        ///     Implements the OnGetImage method in customUI.xml
        /// </summary>
        /// <param name="imageName">the image name in customUI.xml</param>
        /// <returns>memory stream contains image</returns>
        public Bitmap GetImage(IRibbonControl control)
        {
            if (control.Id == "showFormButton")
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("CSOneNoteRibbonAddIn.showform.png"))
                {
                    return new Bitmap(stream);
                }
            }
            return null;
        }

        /// <summary>
        ///     show Windows Form method
        /// </summary>
        /// <param name="control">Represents the object passed into every
        /// Ribbon user interface (UI) control's callback procedure.</param>
        public void ShowForm(IRibbonControl control)
        {
            OneNote.Window context = control.Context as OneNote.Window;
            CWin32WindowWrapper owner =
                new CWin32WindowWrapper((IntPtr)context.WindowHandle);
            TestForm form = new TestForm(applicationObject as OneNote.Application);
            form.ShowDialog(owner);

            form.Dispose();
            form = null;
            context = null;
            owner = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        #endregion
    }

    public class BackupConfig
    {
        public string BackupPath { get; set; }
        public string BackupTime { get; set; } 
        public bool ShouldRun { get; set; }
        public DateTime NextScheduledTime { get; set; }
        public DateTime LastBackupTime { get; set; }

        private static readonly string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BackupConfig.txt");
        public static BackupConfig LoadOrCreate()
        {
            if (!File.Exists(ConfigPath))
            {
                var now = DateTime.Now;
                var config = new BackupConfig
                {
                    BackupPath = @"C:\Backup_Folder",
                    BackupTime = DateTime.Today.ToString("HH:mm"),
                    ShouldRun = true,
                    LastBackupTime = DateTime.MinValue,
                    NextScheduledTime = now
                };
                config.Save();
                return config;
            }

            var lines = File.ReadAllLines(ConfigPath);
            // Format safety
            if (lines.Length != 5)
                throw new InvalidOperationException("BackupConfig.txt is not valid.");

            return new BackupConfig
            {
                BackupPath = lines[0].Trim(),
                BackupTime = lines[1].Trim(),
                ShouldRun = bool.Parse(lines[2].Trim()),
                NextScheduledTime = DateTime.Parse(lines[3].Trim()),
                LastBackupTime = DateTime.Parse(lines[4].Trim())
            };
        }
        public void Save()
        {
            File.WriteAllLines(ConfigPath, new[]
            {
            BackupPath,
            BackupTime,
            ShouldRun.ToString(),
            NextScheduledTime.ToString("o"),
            LastBackupTime.ToString("o")
        });
        }
    }
    public static class AutoExportHelper
    {
        public static async Task CopyFileWithDelayAsync(string backupPath)
        {
            string tablePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bookmarks.txt");
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string destPath = Path.Combine(backupPath, $"OneNote_bookmarks_{timestamp}.txt");

            try
            {
                Directory.CreateDirectory(backupPath);
                if (!File.Exists(tablePath))
                    throw new FileNotFoundException("Source file missing", tablePath);

                await Task.Delay(5000); // Delay before backup

                using (var sourceStream = new FileStream(tablePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var destStream = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await sourceStream.CopyToAsync(destStream);
                }
            }
            catch (Exception ex)
            {
                // Log error as needed
                Console.WriteLine($"Backup failed: {ex}");
            }
        }

        public static async Task RunPeriodicCopyAsync(CancellationToken cancellationToken)
        {
            var config = BackupConfig.LoadOrCreate();

            while (!cancellationToken.IsCancellationRequested)
            {
                var now = DateTime.Now;

                bool missedBackup = now >= config.NextScheduledTime;
                if (config.ShouldRun && missedBackup)
                {
                    await CopyFileWithDelayAsync(config.BackupPath);
                    config.LastBackupTime = now;

                    // Set NextScheduledTime to next day at BackupTime
                    var timeParts = config.BackupTime.Split(':');
                    int hour = int.Parse(timeParts[0]);
                    int minute = int.Parse(timeParts[1]);
                    var nextDay = now.Date.AddDays(1).AddHours(hour).AddMinutes(minute);
                    config.NextScheduledTime = nextDay;
                    config.Save();
                }
                else
                {
                    // Wait until next scheduled time or cancellation
                    var delay = config.NextScheduledTime > now
                        ? config.NextScheduledTime - now
                        : TimeSpan.FromMinutes(1);
                    try
                    {
                        await Task.Delay(delay, cancellationToken);
                    }
                    catch (TaskCanceledException) { }
                }
                // Refresh config in case of external changes
                config = BackupConfig.LoadOrCreate();
            }
        }
    }

}




