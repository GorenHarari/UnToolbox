﻿using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UnToolbox
{
    /// <summary>
    /// Our Solidworks taskpane add-in https://www.youtube.com/watch?v=7DlG6OQeJP0&ab_channel=AngelSix
    /// </summary>
    public class TaskpaneIntegration : ISwAddin
    {
        #region private members
        /// <summary>
        /// The cookie to the cuurent instance of solidworks we are running inside of
        /// </summary>
        private int mSwCookie;
        /// <summary>
        /// The taskpane view for our add-in
        /// </summary>
        private TaskpaneView mTaskpaneView;

        private TaskpaneHostUI mTaskpaneHost;

        private SldWorks mSolidworksApplication;
        // the active assembly
        private ModelDoc2 activeDoc;
        //the path of the selected file, used to get file name
        private string selectedFilePath;
        //the path of the new part
        private string pathMod = string.Empty;
        //DLL location
        static readonly string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase).Replace(@"file:\", "");
        //DLL version
        //private static readonly string version = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().CodeBase).FileVersion;
        //icons location in reference to the DLL
        static readonly string iconPath20x20 = Path.Combine(assemblyFolder , @"Resources\untoolbox icon 20x20.png");
        static readonly string iconPath32x32 = Path.Combine(assemblyFolder , @"Resources\untoolbox icon 32x32.png");
        static readonly string iconPath40x40 = Path.Combine(assemblyFolder , @"Resources\untoolbox icon 40x40.png");
        static readonly string iconPath64x64 = Path.Combine(assemblyFolder , @"Resources\untoolbox icon 64x64.png");
        static readonly string iconPath96x96 = Path.Combine(assemblyFolder , @"Resources\untoolbox icon 96x96.png");
        //icons for the context sensitive menu
        private readonly string[] imageList = {
            iconPath20x20,
            iconPath32x32,
            iconPath40x40,
            iconPath64x64,
            iconPath96x96};
        #endregion

        #region public members

        public const string SWTASKPANE_PROGID = "SimpleM.UntoolBox";

        #endregion

        #region Sollidworks Add-in Callbacks
        /// <summary>
        /// Called when Solidworks has loaded our add-in and wants us to do our conection logic
        /// </summary>
        /// <param name="ThisSW"> the current Solidworks instance</param>
        /// <param name="Cookie">the current Solidworks cookie Id</param>
        /// <returns></returns>
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            mSolidworksApplication = (SldWorks)ThisSW;
            mSwCookie = Cookie;
            Frame swFrame = (Frame)mSolidworksApplication.Frame();

            var ok = mSolidworksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);

            //add the icon to the relevant context sensitive menus
            Debug.WriteLine(assemblyFolder);
            bool resultCodeA = swFrame.AddMenuPopupIcon3((int)swDocumentTypes_e.swDocASSEMBLY, (int)swSelectType_e.swSelFACES, "Third-party context-sensitive", mSwCookie, "CSCallbackFunction", "CSEnable", "", imageList);
            resultCodeA = swFrame.AddMenuPopupIcon3((int)swDocumentTypes_e.swDocASSEMBLY, (int)swSelectType_e.swSelCOMPONENTS, "Third-party context-sensitive", mSwCookie, "CSCallbackFunction", "CSEnable", "", imageList);

            //attach solidworks event listeners
            AttachEventHandlers();

            //task pane UI
            //LoadUI();

            return true;
        }

        public bool DisconnectFromSW()
        {
            //unload taskpane UI
            //UnloadUI();
            return true;
        }
        #endregion

        #region Create TaskPane UI
        private void LoadUI()
        {
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", ""), @"Resources\untoolbox icon_16x16.png");
            mTaskpaneView =  mSolidworksApplication.CreateTaskpaneView2(imagePath,"Untoolbox part");
            mTaskpaneHost = (TaskpaneHostUI)mTaskpaneView.AddControl(TaskpaneIntegration.SWTASKPANE_PROGID, string.Empty);
        }

        private void UnloadUI()
        {
            mTaskpaneHost = null;
            mTaskpaneView.DeleteView();
            Marshal.ReleaseComObject(mTaskpaneView);
            mTaskpaneView = null;
        }

        #endregion

        #region context sensitive UI Callbacks
        //called when the icon in the context sensitive manu is clicked
        public void CSCallbackFunction()
        {
            Debug.WriteLine("Context sensitive menu icon was clicked");

            //get all loade docs names to prevent naming collision with newly created part
            List<string> loadedDocsNames = new List<string>();
            int docCount = mSolidworksApplication.GetDocumentCount();
            Debug.WriteLine(docCount.ToString());
            object[] loadedDocs = mSolidworksApplication.GetDocuments();
            foreach (object doc in loadedDocs)
            {
                ModelDoc2 doc2 = (ModelDoc2)doc;
                string name = Path.GetFileNameWithoutExtension(doc2.GetPathName());
                Debug.WriteLine(name);
                loadedDocsNames.Add(name);
            }


            //initialize the save as dialog box
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.DefaultExt = ".sldprt";
            saveFile.Title = "Save As";
            saveFile.InitialDirectory = Path.GetDirectoryName(activeDoc.GetPathName());
            //store the cuuurent file name to compare to the new file name
            string currentFileName = Path.GetFileName(selectedFilePath);
            saveFile.FileName = Path.GetFileName(currentFileName);

            AssemblyDoc assemblyDoc = (AssemblyDoc)activeDoc;

            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                //make part independent if the save as dialog box is ok
                if (loadedDocsNames.Contains(Path.GetFileNameWithoutExtension(saveFile.FileName)))
                {
                    Debug.WriteLine(saveFile.FileName);
                    MessageBox.Show("File name already exists. \nPlease select different file name");
                    CSCallbackFunction();
                }
                else
                {
                    pathMod = saveFile.FileName;
                    Debug.WriteLine("Save file to: " + pathMod);
                    bool retVal = assemblyDoc.MakeIndependent(pathMod);
                    Debug.WriteLine("Make independent returned: " + retVal.ToString());
                    //add event listener to when solidworks is idle, this is need to implement the referenced file change in the assembly after the make independent
                    mSolidworksApplication.OnIdleNotify += this.swApp_OnIdleNotify;
                }
            }
        }

        //desides wether to show or hide the icon in the context sensitive menu
        public int CSEnable()
        {
            //check wether the selected part in the assembly is toolbox or not
            int isToolboxPart = 0;
            //int retVal = 0;
            //bool isLightWeight = false;
            SelectionMgr selectionMgr = (SelectionMgr)activeDoc.SelectionManager;
            swSelectType_e swSelectType = (swSelectType_e)selectionMgr.GetSelectedObjectType3(1, -1);
            Debug.WriteLine(swSelectType);

            Component2 comp2 = selectionMgr.GetSelectedObjectsComponent4(1, -1);
            //in case root assembly is selected with right click
            if (comp2 == null)
                return 0;

            /// only way to check if part is part is toolbox. resolving lightweight and supressed parts is to slow
            /// 

            // if part is lightweight change to resolved
            //swComponentSuppressionState_e suppressionState = (swComponentSuppressionState_e)comp2.GetSuppression2();
            //if (suppressionState == swComponentSuppressionState_e.swComponentLightweight)
            //{
            //    retVal = comp2.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
            //    isLightWeight = true;
            //}
            //Debug.WriteLine(suppressionState);


            ModelDoc2 selectedModelDoc = comp2.GetModelDoc2();
            //in case component is suppressed or lightweight
            if (selectedModelDoc == null)
                return 0;

            ModelDocExtension modelDocExtension = selectedModelDoc.Extension;
            isToolboxPart = modelDocExtension.ToolboxPartType;

            //return part to lightweight state
            //if(isLightWeight)
            //{
            //    retVal = comp2.SetSuppression2((int)swComponentSuppressionState_e.swComponentLightweight);
            //}

            //if it is toolbox show the icon in the context sensitive menu, hide it if it not
            if (isToolboxPart != (int)swToolBoxPartType_e.swNotAToolboxPart)
            {
                Debug.WriteLine("Toolbox part was selected in a assembly document.");
                selectedFilePath = selectedModelDoc.GetPathName();
                return 1;
            }

            return 0;
        }
        #endregion

        #region events handlers

        public void AttachEventHandlers()
        {
            AttachSWEvents();
        }


        private void AttachSWEvents()
        {
            mSolidworksApplication.ActiveDocChangeNotify += this.swApp_ActiveDocChangeNotify;
        }

        //get the model doc of the active doc
        private int swApp_ActiveDocChangeNotify()
        {
            Debug.WriteLine("Active Doc changed event fired");
            activeDoc = mSolidworksApplication.ActiveDoc;
            return 0;
        }

        //solidworks idle event listener
        private int swApp_OnIdleNotify()
        {
            //get the path of the newly created file to untoolbox it
            Debug.WriteLine("On idle event fired");
            AssemblyDoc assemblyDoc = (AssemblyDoc)activeDoc;
            string fileToModify = Path.GetFileNameWithoutExtension(pathMod);
            int count = assemblyDoc.GetComponentCount(false);
            Debug.WriteLine("amount of components including children: " + count.ToString());
            object[] components = assemblyDoc.GetComponents(false);
            foreach(object compInstance in components)
            {
                Component2 component = (Component2)compInstance;
                string compName = component.Name2;
                Debug.WriteLine(compName);
            }

            //untoolbox the part
            Component2 comp = assemblyDoc.GetComponentByName(fileToModify+"-1");
            ModelDoc2 newModelDoc2 = comp.GetModelDoc2();
            ModelDocExtension modelDocExtension = newModelDoc2.Extension;
            modelDocExtension.ToolboxPartType = (int)swToolBoxPartType_e.swNotAToolboxPart;
            

            int lErrors=0;
            int lWarnings=0;
            
            //save the part and the assembly
            bool status = newModelDoc2.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);
            Debug.WriteLine("save was successful: " + status.ToString());
            //activeDoc.ForceRebuild3(false);
            //modelDocExtension = activeDoc.Extension;
            //status = activeDoc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);
            //modelDocExtension.Rebuild((int)swRebuildOptions_e.swRebuildAll);
            //Debug.WriteLine("save was successful: " + status.ToString());

            //remove event listener to avoid running code each tome solidworks becomes idle
            mSolidworksApplication.OnIdleNotify -= this.swApp_OnIdleNotify;
            mSolidworksApplication.OnIdleNotify += this.swApp_OnIdleNotify2;

            return 0;
        }

        private int swApp_OnIdleNotify2()
        {
            int lErrors = 0;
            int lWarnings = 0;

            //ModelDocExtension modelDocExtension = activeDoc.Extension;
            //modelDocExtension.Rebuild((int)swRebuildOptions_e.swForceRebuildAll);
            activeDoc.ForceRebuild3(false);
            bool status = activeDoc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);
            Debug.WriteLine("save was successful: " + status.ToString());

            mSolidworksApplication.OnIdleNotify -= this.swApp_OnIdleNotify2;
            return 0;
        }

        #endregion

        #region COM registration
        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            var ketPath = string.Format(@"SOFTWARE\Solidworks\AddIns\{0:b}", t.GUID);
            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(ketPath))
            {
                //Load Add-in when solidworks opens
                rk.SetValue(null, 1);

                //set solidworks add-in title and description
                rk.SetValue("Title", "UnToolbox");
                rk.SetValue("Description", "convert toolbox part to normal part");
            }
        }

        [ComUnregisterFunction]
        private static void ComUnRegister(Type t)
        {
            var ketPath = string.Format(@"SOFTWARE\Solidworks\AddIns\{0:b}", t.GUID);
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(ketPath);
        }

        #endregion
    }
}
