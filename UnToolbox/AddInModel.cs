using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.swcommands;
using System.Windows.Forms;

namespace UnToolbox
{
    class AddInModel
    {
        SldWorks swApp;
        ModelDoc2 swModel;
        PartDoc pDoc;
        AssemblyDoc aDoc;
        DrawingDoc dDoc;
        int isToolboxPart;

        public AddInModel(SldWorks _sldwrks, int addinID)
        {
            swApp = _sldwrks;
            Frame swFrame = swApp.Frame();
            string[] imageList = new string[6];
            imageList[0] = @"C:\Users\Goren Harari\source\repos\UnToolbox\UnToolbox\bin\Debug\untoolbox icon 20x20.png";
            imageList[1] = @"C:\Users\Goren Harari\source\repos\UnToolbox\UnToolbox\bin\Debug\untoolbox icon 32x32.png";
            imageList[2] = @"C:\Users\Goren Harari\source\repos\UnToolbox\UnToolbox\bin\Debug\untoolbox icon 40x40.png";
            imageList[3] = @"C:\Users\Goren Harari\source\repos\UnToolbox\UnToolbox\bin\Debug\untoolbox icon 64x64.png";
            imageList[4] = @"C:\Users\Goren Harari\source\repos\UnToolbox\UnToolbox\bin\Debug\untoolbox icon 96x96.png";
            imageList[5] = @"C:\Users\Goren Harari\source\repos\UnToolbox\UnToolbox\bin\Debug\untoolbox icon 128x128.png";
            bool resultCode = swFrame.AddMenuPopupIcon3((int)swDocumentTypes_e.swDocPART, (int)swSelectType_e.swSelFACES, "Third-party context-sensitive", addinID, "", "", "", imageList);

            //swModel = (ModelDoc2)swApp.ActiveDoc;

            //if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            //{
            //    aDoc = (AssemblyDoc)swModel;
            //}
            //else if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            //{
            //    pDoc = (PartDoc)swModel;
            //}
            //else if (swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
            //{
            //    dDoc = (DrawingDoc)swModel;
            //}

            AttachEventHandlers();
        }

        public void AttachEventHandlers()
        {
            AttachSWEvents();
        }


        private void AttachSWEvents()
        {
            swApp.ActiveDocChangeNotify += this.swApp_ActiveDocChangeNotify;

            //if ((aDoc != null))
            //{
            //    aDoc.UserSelectionPostNotify += this.aDoc_UserSelectionPostNotify;
            //}
            //if ((pDoc != null))
            //{
            //    pDoc.UserSelectionPostNotify += this.pDoc_UserSelectionPostNotify;
            //}
            //if ((dDoc != null))
            //{
            //    dDoc.UserSelectionPostNotify += this.dDoc_UserSelectionPostNotify;
            //}
        }

        private void AttachModelDocEvents(ModelDoc2 mDoc)
        {
            SelectionMgr selectionMgr = (SelectionMgr)mDoc.SelectionManager;
            if (mDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                aDoc = (AssemblyDoc)mDoc;
                aDoc.UserSelectionPostNotify += this.aDoc_UserSelectionPostNotify;
            }
            //else if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            //{
            //    pDoc = (PartDoc)swModel;
            //}
            //else if (swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
            //{
            //    dDoc = (DrawingDoc)swModel;
            //}
        }

        private int swApp_ActiveDocChangeNotify()
        {
            
            Debug.WriteLine("Active Doc changed event fired");
            swModel = swApp.ActiveDoc;
            AttachModelDocEvents(swModel);
            return 0;
        }

        public int aDoc_UserSelectionPostNotify()
        {
            int functionReturnValue = 0;
            SelectionMgr selectionMgr = (SelectionMgr)swModel.SelectionManager;
            Component2 comp2 = selectionMgr.GetSelectedObjectsComponent4(1, -1);
            ModelDoc2 selectedModelDoc = comp2.GetModelDoc2();
            ModelDocExtension modelDocExtension = selectedModelDoc.Extension;
            isToolboxPart = modelDocExtension.ToolboxPartType;
            if(isToolboxPart != (int)swToolBoxPartType_e.swNotAToolboxPart)
                Debug.WriteLine("Toolbox part was selected in a assembly document.");

            return functionReturnValue;
        }

        //private int pDoc_UserSelectionPostNotify()
        //{
        //    int functionReturnValue = 0;
        //    MessageBox.Show("An entity was selected in a part document.");
        //    return functionReturnValue;
        //}

        //private int dDoc_UserSelectionPostNotify()
        //{
        //    int functionReturnValue = 0;
        //    MessageBox.Show("An entity was selected in a drawing document.");
        //    return functionReturnValue;
        //}
    }
}
