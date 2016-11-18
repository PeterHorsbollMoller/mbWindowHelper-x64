using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MapInfo.Types;
using System.Windows.Forms;

namespace WindowHelper
{
    class ZoomNextAndPrevious
    {
        //*****************************************************************************    
        #region private variables
//        private static IMapInfoPro _mapinfoApp;
//        private static bool _MapInfoInitialised = false;
        private static Windows.Window _activeWindow = null;

        private static Windows.MapWindowExtentsList _mapWindowExtents = new Windows.MapWindowExtentsList();

        #endregion
 
        //*****************************************************************************    
        //#region Constructors
        //public ZoomNextAndPrevious()
        //{
        //    _miApp = InteropServices.MapInfoApplication;
        //    //_miApp.Do("Set Paper Units \"cm\"");
        //    // Register the window with the docking system
        //}
        //#endregion 

        //*****************************************************************************    
        #region Mehods
        //-----------------------------------------------------------------
/*        public static void Initialise(IMapInfoPro mapinfoApp)
        {
            if (!_MapInfoInitialised)
            {
                _mapinfoApp = mapinfoApp;
                _MapInfoInitialised = true;
            }
        }
*/
        //-----------------------------------------------------------------
        public static void ResetActiveWindow()
        {
            if (_activeWindow != null)
                _activeWindow = null;
        }

        //-----------------------------------------------------------------
        public static void SetActiveWindow(int windowID)
        {
            Windows.Window window = new Windows.Window(windowID);
            _activeWindow = window;
        }

        #endregion

        //-----------------------------------------------------------------
        public static void AddWindow(int windowID)
        {
            try
            {
                Windows.Window mapWindow = new Windows.Window(windowID);
                if (mapWindow.Type != Windows.Window.WindowType.Mapper)
                    return;

                if (_mapWindowExtents.WindowExists(mapWindow) == false)
                {
                    Windows.MapWindowExtents mapExtent = new Windows.MapWindowExtents(mapWindow);
                    _mapWindowExtents.Add(mapExtent);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("{0} Exception caught.", e));
            }
        }

        //-----------------------------------------------------------------
        public static void RemoveWindow(int windowID)
        {
            Windows.Window mapWindow = new Windows.Window(windowID);
            if (mapWindow.Type != Windows.Window.WindowType.Mapper)
                return;

            int index = _mapWindowExtents.FindWindow(mapWindow);
            if (index >= 0)
                _mapWindowExtents.RemoveAt(index);
        }

        //-----------------------------------------------------------------
        public static void AddExtent(int windowID)
        {
            try
            {
                Windows.Window mapWindow = new Windows.Window(windowID);
                if (mapWindow.Type != Windows.Window.WindowType.Mapper)
                    return;

                int index = _mapWindowExtents.FindWindow(mapWindow);
                if (index >= 0)
                {
                    int extentIndex = _mapWindowExtents[index].MapExtentList.AddExtent(mapWindow);
                }
                else
                {
                    AddWindow(windowID);
                }
            }
            catch (Exception e)
            {
                WindowHelper.InteropHelper.PrintMessage(string.Format("{0} Exception caught.", e));
            }
        }

        //-----------------------------------------------------------------
        public static void ZoomToPrevious()
        {
            if (_activeWindow.Type != Windows.Window.WindowType.Mapper)
                return;

            int index = _mapWindowExtents.FindWindow(_activeWindow);
            if (index >= 0)
            {
                _mapWindowExtents[index].MapExtentList.ZoomToPrevious(_activeWindow);
            }
        }
        //-----------------------------------------------------------------
        public static void ZoomToNext()
        {
            if (_activeWindow.Type != Windows.Window.WindowType.Mapper)
                return;

            int index = _mapWindowExtents.FindWindow(_activeWindow);
            if (index >= 0)
            {
                _mapWindowExtents[index].MapExtentList.ZoomToNext(_activeWindow);
            }
        }
        //-----------------------------------------------------------------
        public static void ZoomToFirst()
        {
            if (_activeWindow.Type != Windows.Window.WindowType.Mapper)
                return;

            int index = _mapWindowExtents.FindWindow(_activeWindow);
            if (index >= 0)
            {
                _mapWindowExtents[index].MapExtentList.ZoomToFirst(_activeWindow);
            }
        }
        //-----------------------------------------------------------------
        public static void ZoomToLast()
        {
            if (_activeWindow.Type != Windows.Window.WindowType.Mapper)
                return;

            int index = _mapWindowExtents.FindWindow(_activeWindow);
            if (index >= 0)
            {
                _mapWindowExtents[index].MapExtentList.ZoomToLast(_activeWindow);
            }
        }
    }
}
