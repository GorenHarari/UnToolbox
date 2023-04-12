using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
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

            //initialize the save as dialob box
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.DefaultExt = ".sldprt";
            saveFile.AddExtension = true;
            saveFile.Title = "Save As";
            saveFile.InitialDirectory = Path.GetDirectoryName(activeDoc.GetPathName());
            //store the cuuurent file name to compare to the new file name
            string currentFileName = Path.GetFileName(selectedFilePath);
            saveFile.FileName = Path.GetFileName(currentFileName);

            AssemblyDoc assemblyDoc = (AssemblyDoc)activeDoc;
            
            //make part independent if the save as dialog box is ok
            if (saveFile.ShowDialog() == DialogResult.OK && saveFile.FileName != currentFileName)
            {
                pathMod = saveFile.FileName;
                Debug.WriteLine("Save file to: " + pathMod);
                bool retVal = assemblyDoc.MakeIndependent(pathMod);
                Debug.WriteLine("Make independent returned: " + retVal.ToString());
            }
            else
            {
                MessageBox.Show("please select different file name");
                CSCallbackFunction();
            }

            //add event listener to when solidworks is idle, this is need to implement the referenced file change in the assembly after the make independent
            mSolidworksApplication.OnIdleNotify += this.swApp_OnIdleNotify;
        }

        //desides wether to show or hide the icon in the context sensitive menu
        public int CSEnable()
        {
            //check wether the selected part in the assembly is toolbox or not
            int isToolboxPart = 0;
            SelectionMgr selectionMgr = (SelectionMgr)activeDoc.SelectionManager;
            Component2 comp2 = selectionMgr.GetSelectedObjectsComponent4(1, -1);
            ModelDoc2 selectedModelDoc = comp2.GetModelDoc2();
            ModelDocExtension modelDocExtension = selectedModelDoc.Extension;
            isToolboxPart = modelDocExtension.ToolboxPartType;
            //if it is toolbox show the icon in the context sensitive menu, hide it if it not
            if (isToolboxPart != (int)swToolBoxPartType_e.swNotAToolboxPart)
            {
                Debug.WriteLine("Toolbox part was selected in a assembly document.");
                selectedFilePath = selectedModelDoc.GetPathName();
                isToolboxPart = 1;
            }
            return isToolboxPart;
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
            mSolidworksApplication.OnIdleNotify -= this.swApp_OnIdleNotify;

            int lErrors=0;
            int lWarnings=0;
            
            //save the part and the assembly
            bool status = newModelDoc2.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);
            Debug.WriteLine("save was successful: " + status.ToString());
            activeDoc.ForceRebuild3(false);
            status = activeDoc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref lErrors, ref lWarnings);
            Debug.WriteLine("save was successful: " + status.ToString());
            //activeDoc.ForceRebuild3(false);

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
                rk.SetValue(null, 0);

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
