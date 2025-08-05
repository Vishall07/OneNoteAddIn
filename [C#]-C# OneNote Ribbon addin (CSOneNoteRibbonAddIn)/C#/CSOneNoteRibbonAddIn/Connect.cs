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
    using System;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Windows.Forms;
    using System.Xml.Linq;
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

        private object applicationObject;
        private object addInInstance;

        /// <summary>
        ///     Loads the XML markup from an XML customization file 
        ///     that customizes the Ribbon user interface.
        /// </summary>
        /// <param name="RibbonID">The ID for the RibbonX UI</param>
        /// <returns>string</returns>
        public string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
              <ribbon>
                <tabs>
                  <tab idMso='TabHome'>
                    <group id='customGroup' label='BookMark'>
                      <button id='showFormButton'
                              label='Save'
                              imageMso='BookMark'
                              size='large'
                              onAction='OnShowFormButtonClick' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }
        private IRibbonUI ribbon;

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        // This method name must match the XML "onAction" value!
        public void OnShowFormButtonClick(IRibbonControl control)
        {
            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                try
                {
                    System.Windows.Forms.Application.EnableVisualStyles();
                    System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

                    var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
                    string hierarchyXml;
                    oneNoteApp.GetHierarchy("", Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, out hierarchyXml);

                    // Here: Determine current selection - simplified for example
                    string selectedId = null;
                    string selectedScope = "Paragraph"; // default scope

                    // Try to get current selection; if null, select first para id
                    try
                    {
                        // Get current page id
                        var window = oneNoteApp.Windows.CurrentWindow;
                        string currentPageId = window.CurrentPageId;

                        string pageXml;
                        oneNoteApp.GetPageContent(currentPageId, out pageXml, Microsoft.Office.Interop.OneNote.PageInfo.piAll);

                        var doc = new System.Xml.XmlDocument();
                        doc.LoadXml(pageXml);
                        var nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                        nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                        // Try get selected paragraph ID, if fail get first paragraph node id
                        var selectedOutline = doc.SelectSingleNode("//one:Outline[@selected='true']", nsmgr);
                        if (selectedOutline != null)
                        {
                            selectedId = selectedOutline.Attributes["ID"]?.Value;
                        }
                        else
                        {
                            // Select first paragraph outline ID
                            var firstOutline = doc.SelectSingleNode("//one:Outline", nsmgr);
                            selectedId = firstOutline?.Attributes["ID"]?.Value;
                        }

                        if (string.IsNullOrEmpty(selectedId))
                        {
                            // fallback to current page id, scope page
                            selectedId = currentPageId;
                            selectedScope = "Page";
                        }
                    }
                    catch
                    {
                        // default fallback
                        selectedId = null;
                        selectedScope = "Page";
                    }

                    string displayText = "Current Selection"; // Optionally get a name for UI

                    var form = new BookMark_Window(selectedId, selectedScope, displayText);
                    System.Windows.Forms.Application.Run(form);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Error launching bookmark window: " + ex.Message);
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
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