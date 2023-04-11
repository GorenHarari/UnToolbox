using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

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

            var ok = mSolidworksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);
            AddInModel addInModel = new AddInModel(mSolidworksApplication, Cookie);

            LoadUI();

            return true;
        }

        public bool DisconnectFromSW()
        {
            UnloadUI();
            return true;
        }
        #endregion

        #region Create UI
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
                rk.SetValue("Description", "convert toolbox part tonormal part");
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
