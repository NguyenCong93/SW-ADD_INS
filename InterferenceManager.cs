using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomationAddin
{
    public class InterferenceManager
    {
        private ISldWorks _swApp;
        private InterferenceDetectionMgr _activeInterferenceMgr;
        private List<IInterference> _currentClashes;
        private int _currentClashIndex = -1;

        public InterferenceManager(ISldWorks swApp)
        {
            _swApp = swApp;
            _currentClashes = new List<IInterference>();
        }

        public int RunInterferenceDetectionAndHighlight()
        {
            try
            {
                if (_activeInterferenceMgr != null)
                {
                    _activeInterferenceMgr.Done();
                    _activeInterferenceMgr = null;
                }

                _currentClashes.Clear();
                _currentClashIndex = -1;

                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                    return -1;

                AssemblyDoc swAssy = (AssemblyDoc)swModel;

                // Tự động chuyển hình chiếu trục đo và zoom vừa vặn
                swModel.ShowNamedView2("*Isometric", 7);
                swModel.ViewZoomtofit2();

                _activeInterferenceMgr = swAssy.InterferenceDetectionManager;
                if (_activeInterferenceMgr == null) return -1;

                _activeInterferenceMgr.TreatCoincidenceAsInterference = false;
                _activeInterferenceMgr.ShowIgnoredInterferences = false;
                _activeInterferenceMgr.IgnoreHiddenBodies = true;
                _activeInterferenceMgr.CreateFastenersFolder = true;
                // Làm trong các chi tiết không va chạm (Transparent)
                _activeInterferenceMgr.NonInterferingComponentDisplay = 2;

                swModel.ClearSelection2(true);
                int interferenceCount = _activeInterferenceMgr.GetInterferenceCount();

                object[] interferences = (object[])_activeInterferenceMgr.GetInterferences();
                if (interferences == null || interferences.Length == 0)
                {
                    _activeInterferenceMgr.Done();
                    _activeInterferenceMgr = null;
                    return 0;
                }

                int actualClashCount = 0;
                foreach (object obj in interferences)
                {
                    IInterference swInterf = (IInterference)obj;
                    if (swInterf.IsFastener)
                    {
                        swInterf.Ignore = true;
                        continue;
                    }

                    actualClashCount++;
                    _currentClashes.Add(swInterf);
                }

                if (actualClashCount == 0)
                {
                    _activeInterferenceMgr.Done();
                    _activeInterferenceMgr = null;
                }

                return actualClashCount;
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi trong quá trình kiểm tra va chạm: " + ex.Message);
                return -1;
            }
        }

        public void ZoomToNextClash()
        {
            if (_currentClashes == null || _currentClashes.Count == 0) return;
            
            _currentClashIndex++;
            if (_currentClashIndex >= _currentClashes.Count)
            {
                _currentClashIndex = 0; 
            }
            ZoomToCurrentIndex();
        }

        public void ZoomToPrevClash()
        {
            if (_currentClashes == null || _currentClashes.Count == 0) return;
            
            _currentClashIndex--;
            if (_currentClashIndex < 0)
            {
                _currentClashIndex = _currentClashes.Count - 1; 
            }
            ZoomToCurrentIndex();
        }

        private void ZoomToCurrentIndex()
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel == null || _activeInterferenceMgr == null) return;

                swModel.ClearSelection2(true);
                IInterference current = _currentClashes[_currentClashIndex];

                // Highlight components
                object[] components = (object[])current.Components;
                if (components != null)
                {
                    foreach (object comp in components)
                    {
                        Component2 swComp = (Component2)comp;
                        swComp.Select4(true, null, false);
                    }
                }

                // Get exact interference body box and zoom to it
                Body2 interBody = (Body2)current.GetInterferenceBody();
                if (interBody != null)
                {
                    object boxObj = interBody.GetBodyBox();
                    if (boxObj != null)
                    {
                        double[] box = (double[])boxObj;
                        if (box.Length >= 6)
                        {
                            // Zoom out slightly (0.5x) for tighter context (zoom 2x closer than previous 1.5x)
                            double dx = (box[3] - box[0]) * 0.5;
                            double dy = (box[4] - box[1]) * 0.5;
                            double dz = (box[5] - box[2]) * 0.5;
                            // Make sure dx, dy, dz aren't zero
                            if (dx < 0.01) dx = 0.05;
                            if (dy < 0.01) dy = 0.05;
                            if (dz < 0.01) dz = 0.05;

                            swModel.ViewZoomTo2(box[0] - dx, box[1] - dy, box[2] - dz, box[3] + dx, box[4] + dy, box[5] + dz);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi zoom: " + ex.Message);
            }
        }

        public void EndClashDetection()
        {
            try
            {
                if (_activeInterferenceMgr != null)
                {
                    _activeInterferenceMgr.Done();
                    _activeInterferenceMgr = null;
                }
                _currentClashes.Clear();
                _currentClashIndex = -1;

                ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;
                if (swModel != null)
                {
                    swModel.ClearSelection2(true);
                    swModel.ShowNamedView2("*Front", 1);
                    swModel.ViewZoomtofit2();
                }
            }
            catch (Exception ex)
            {
                _swApp.SendMsgToUser("Lỗi khi kết thúc: " + ex.Message);
            }
        }
    }
}
