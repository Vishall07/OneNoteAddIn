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
    using System.Threading;
    using System.Windows.Forms;
    using System.Xml;
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
                var oneNoteApp = new OneNote.Application();
                var model = GetCurrentNotebookModel(oneNoteApp);
                if (model == null)
                {
                    MessageBox.Show("Failed to load the current notebook model.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

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
            try
            {
                string pageXml;
                oneNoteApp.GetPageContent(page.Id, out pageXml, PageInfo.piAll);

                var doc = new System.Xml.XmlDocument();
                doc.LoadXml(pageXml);

                var nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                // Select all paragraphs inside page content (Text elements inside Outline > OEChildren > OE)
                // This XPath targets the text content inside paragraph nodes
                var paragraphNodes = doc.SelectNodes("//one:Outline/one:OEChildren/one:OE/one:T", nsmgr);
                if (paragraphNodes != null)
                {
                    int index = 1;
                    foreach (System.Xml.XmlNode paraNode in paragraphNodes)
                    {
                        string paraText = paraNode.InnerText?.Trim();
                        if (!string.IsNullOrEmpty(paraText))
                        {
                            page.Paragraphs.Add(new ParagraphModel
                            {
                                // There is no ID on text nodes, so an index based Id or other unique ID can be used here
                                Id = page.Id + "_para_" + index,
                                Name = paraText
                            });
                            index++;
                        }
                        break;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading paragraphs for page {page.Name}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
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

            var cursorPos = Cursor.Position;
            var screen = Screen.FromPoint(cursorPos);
            bool goesOffRight = cursorPos.X + form.Width > screen.WorkingArea.Right;

            int x = goesOffRight
                ? cursorPos.X - form.Width
                : cursorPos.X;

            int y = cursorPos.Y;

            if (y < screen.WorkingArea.Top)
                y = screen.WorkingArea.Top;
            if (y + form.Height > screen.WorkingArea.Bottom)
                y = screen.WorkingArea.Bottom - form.Height;

            form.Left = x;
            form.Top = y;
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
}