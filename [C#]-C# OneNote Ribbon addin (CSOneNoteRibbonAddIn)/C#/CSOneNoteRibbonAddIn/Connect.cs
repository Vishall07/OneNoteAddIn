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
    using CSOneNoteRibbonAddIn.Properties;
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.OneNote;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Threading;
    using System.Windows.Forms;
    using System.Xml.Linq;
    using static System.Net.Mime.MediaTypeNames;
    using static System.Windows.Forms.VisualStyles.VisualStyleElement;
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
        private Thread _uiThread;
        private BookMark_Window _bookmarkWindow;
        private OneNote.Application _oneNoteApp;

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
            _uiThread.Abort();
            _bookmarkWindow.Close();
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

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        public void OnShowFormButtonClick(IRibbonControl control)
        {
            try
            {
                _oneNoteApp = new OneNote.Application();
                var model = GetCurrentNotebookModel(_oneNoteApp);
                if (model == null)
                {
                    MessageBox.Show("Failed to load the current notebook model.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Initialize variables from model
                string selectedId = model.Page?.Id ?? "";
                string selectedScope = "page"; // or whatever your scope is
                string displayText = model.Page?.Name ?? "";

                string notebookName = model.NotebookName ?? "";
                string notebookColor = model.NotebookColor ?? "";

                string sectionGroupName = model.SectionGroup?.Name ?? "";
                string sectionName = model.Section?.Name ?? "";
                string sectionColor = model.Section?.Color ?? "";

                string pageName = model.Page?.Name ?? "";
                string paraContent = model.Page?.Paragraphs?.FirstOrDefault()?.Name ?? "";

                // If the form and thread are already created and alive
                if (_uiThread != null && _uiThread.IsAlive && _bookmarkWindow != null)
                {
                    _bookmarkWindow.Invoke((Action)(() =>
                    {
                        _bookmarkWindow.UpdateBookmarkInfo(
                            selectedId, selectedScope, displayText,
                            notebookName, notebookColor,
                            sectionGroupName, sectionName, sectionColor,
                            pageName, paraContent);
                        _bookmarkWindow.Activate();
                    }));
                    return;
                }

                // If first time: start UI thread and create form
                _uiThread = new Thread(() =>
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
                        System.Windows.Forms.MessageBox.Show("Error launching bookmark window: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });

                _uiThread.SetApartmentState(ApartmentState.STA);
                _uiThread.IsBackground = true;
                _uiThread.Start();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Unexpected error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


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

        public void GetSelectedOutlineInfo(
            OneNote.Application oneNoteApp,
            out string selectedId, out string selectedScope,
            out string notebookName, out string notebookColor,
            out string sectionGroupName, out string sectionName, out string sectionColor,
            out string pageName, out string paraContent)
        {
            // Initialize all out parameters
            selectedId = selectedScope = notebookName = notebookColor = sectionGroupName = sectionName = sectionColor = pageName = paraContent = null;
            selectedScope = "Paragraph"; // default scope

            try
            {
                // 1. Get the full hierarchy XML (contains metadata about notebooks, sections, pages)
                string hierarchyXml;
                oneNoteApp.GetHierarchy("", HierarchyScope.hsPages, out hierarchyXml);
                var hierarchyDoc = new System.Xml.XmlDocument();
                hierarchyDoc.LoadXml(hierarchyXml);

                // Extract the namespace URI dynamically for hierarchy XML
                var hierarchyNsUri = hierarchyDoc.DocumentElement.NamespaceURI;
                var nsmgrHierarchy = new System.Xml.XmlNamespaceManager(hierarchyDoc.NameTable);
                nsmgrHierarchy.AddNamespace("one", hierarchyNsUri);

                // 2. Get current window and page ID
                var window = oneNoteApp.Windows.CurrentWindow;
                string currentPageId = window.CurrentPageId;

                // 3. Get the content XML of the current page
                string pageXml;
                oneNoteApp.GetPageContent(currentPageId, out pageXml, PageInfo.piAll);
                var pageDoc = new System.Xml.XmlDocument();
                pageDoc.LoadXml(pageXml);

                // Extract the namespace URI dynamically for page XML
                var pageNsUri = pageDoc.DocumentElement.NamespaceURI;
                var nsmgrPage = new System.Xml.XmlNamespaceManager(pageDoc.NameTable);
                nsmgrPage.AddNamespace("one", pageNsUri);

                // 4. Extract page name from the page XML
                var pageNode = pageDoc.SelectSingleNode("//one:Page", nsmgrPage);
                pageName = pageNode?.Attributes["name"]?.Value;

                // 5. Find selected outline node or fallback to first outline node
                var selectedOutline = pageDoc.SelectSingleNode("//one:Outline[@selected='true']", nsmgrPage)
                                     ?? pageDoc.SelectSingleNode("//one:Outline", nsmgrPage);

                if (selectedOutline != null)
                {
                    selectedId = selectedOutline.Attributes["ID"]?.Value;
                }
                else
                {
                    // If no outline found, fallback to page ID and scope Page
                    selectedId = currentPageId;
                    selectedScope = "Page";
                }

                // 6. Extract paragraph content from outline (Text inside OE > T nodes)
                var paraNode = selectedOutline?.SelectSingleNode(".//one:OE/one:T", nsmgrPage);
                paraContent = paraNode?.InnerText;

                // 7. Find the page node in the hierarchy XML (to get the section and notebook metadata)
                var pageNodeInHierarchy = hierarchyDoc.SelectSingleNode($"//one:Page[@ID='{currentPageId}']", nsmgrHierarchy);

                var sectionNode = pageNodeInHierarchy?.ParentNode;
                sectionName = sectionNode?.Attributes["name"]?.Value;
                sectionColor = sectionNode?.Attributes["color"]?.Value;
                sectionGroupName = sectionNode?.ParentNode?.Attributes["name"]?.Value;

                // Traverse up to find notebook node
                var notebookNode = sectionNode;
                while (notebookNode != null && notebookNode.Name != "one:Notebook")
                {
                    notebookNode = notebookNode.ParentNode;
                }
                notebookName = notebookNode?.Attributes["name"]?.Value;
                notebookColor = notebookNode?.Attributes["color"]?.Value;
            }
            catch (Exception ex)
            {
                // Reset all outputs on exception
                selectedId = null;
                selectedScope = "Page";
                notebookName = notebookColor = sectionGroupName = sectionName = sectionColor = pageName = paraContent = null;

                // Optionally log the exception message here for diagnostics:
                // Console.WriteLine("Exception in GetSelectedOutlineInfo: " + ex.Message);
            }
        }

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

        public void PositionFormNearCursor(Form form)
        {
            var cursorPos = Cursor.Position;

            int x = cursorPos.X - (form.Width / 2);
            int y = cursorPos.Y + 40;

            var screen = Screen.FromPoint(cursorPos);

            if (x < screen.WorkingArea.Left)
                x = screen.WorkingArea.Left;
            if ((x + form.Width) > screen.WorkingArea.Right)
                x = screen.WorkingArea.Right - form.Width;
            if ((y + form.Height) > screen.WorkingArea.Bottom)
                y = screen.WorkingArea.Bottom - form.Height;

            form.Left = x;
            form.Top = y;
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


    }
}