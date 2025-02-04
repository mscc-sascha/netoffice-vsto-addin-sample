using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools;
using NetOffice.OfficeApi;
using Word02AddinCS4.Properties;
using System.Diagnostics.Contracts;
using System.Linq;
using NetOffice.WordApi;
/*
    Ribbons & Panes Addin Example
*/
namespace Word02AddinCS4
{
    [COMAddin("Word02 Sample Addin CS4", "Ribbons & Panes Addin Example", 3)]
    [ProgId("Word02AddinCS4.Connect"), Guid("E7E8652F-7F9C-48E5-BC7A-7CD5375057AB")]
    [CustomUI("RibbonUI.xml", true)]
    [CustomPane(typeof(SampleUserControl), "Simple Taskpane", false, PaneDockPosition.msoCTPDockPositionRight)]
    [ComVisible(true)]
    public class Addin : COMAddin
    {
        public Addin()
        {
        }

        // Taskpane visibility has been changed. We upate the checkbutton in the ribbon ui for show/hide taskpane
        protected override void TaskPaneVisibleStateChanged(Office._CustomTaskPane customTaskPaneInst)
        {
            if (null != RibbonUI)
                RibbonUI.InvalidateControl("PaneVisibleToogleButton");
        }

        // Defined in RibbonUI.xml to make sure the checkbutton state is up-to-date and synchronized with taskpane visibility.
        public bool OnGetPressedPanelToggle(Office.IRibbonControl control)
        {
            if (TaskPanes.Count > 0)
                return TaskPanes[0].Visible;
            else
                return false;
        }

        // Defined in RibbonUI.xml to track the user clicked ouer checkbutton. Then we upate the panel visibility at hand.
        public void OnCheckPanelToggle(Office.IRibbonControl control, bool pressed)
        {
            if (TaskPanes.Count > 0)
                TaskPanes[0].Visible = pressed;
        }
    }
}
