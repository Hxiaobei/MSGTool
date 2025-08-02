using SolidWorks.Interop.sldworks;
using Sld = SolidWorks.Interop.sldworks;
using swView = SolidWorks.Interop.sldworks.View;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdocumentmgr;
using System;
using System.Windows.Forms;
using System.Linq;
//using SldWorksEx;
namespace SldWorksDemo
{
    public partial class Form1 : Form
    {
        static SldWorks swApp;
        public Form1()
        {
            InitializeComponent();
            if(swApp == null) swApp = SWUtils.GetSWApp();
            if(swApp == null) this.Close();
        }
        void GetDocType()
        {
            msgBox.AppendText($"{(swDocumentTypes_e)swApp.ActiveDoc.GetType()}");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //var swKey = SWUtils.SwDocumentManagerLicenseKey();
            //msgBox.AppendText($"\r\n{swKey}");
            var swModel = swApp.ActiveDoc as ModelDoc2;
            msgBox.AppendText($"标题:{swModel.GetTitle()},{swModel.GetType()}");
            msgBox.AppendText($"\r\n选择对象数量:{SwDemoEx.SelectMgr.GetSelectedObjectCount2(-1)}");
            //msgBox.AppendText($"\r\n选择对象类型:{SwDemoEx.SelectType}");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(!(swApp.ActiveDoc is DrawingDoc drawDoc)) return;
            var sheetNames = drawDoc.GetSheetNames() as string[];
            foreach(string sheetName in sheetNames)
            {
                //drawDoc.ActivateSheet(sheetName);
                Sld.Sheet swSheet = drawDoc.Sheet[sheetName];

                if(!(swSheet.GetViews() is object[] swViews)) continue;

                foreach(swView swView in swViews.Cast<swView>())
                {
                    if(!swView.IsFlatPatternView()) continue;

                    msgBox.AppendText($"\r\n{swView.GetName2()}");
                    msgBox.AppendText($"\r\n{swView.GetBendLineCount()}");
                    msgBox.AppendText($"\r\n{swView.GetPolyLinesAndCurvesCount(0,out var pointCount)}");

                    var bendLines = swView.GetBendLines() as object[];

                    var arcCount = swView.GetArcCount();
                    var arcs = swView.GetArcs4() as object[];

                    var lines = swView.GetLines3();
                    var spline = swView.GetSplines3();
                    //var plyLines = swView.GetPolylines7();
                    //GetPolyLinesAndCurves
                    //获取所有曲线
                }
            }
            //var swSheet = drawDoc.GetCurrentSheet() as Sheet; 获取活动图页
        }
    }
    public static class SwDemoEx
    {
        public static SldWorks SWApp => SWUtils.GetSWApp();
        public static ModelDoc2 ActiveModel => SWApp.ActiveDoc as ModelDoc2;
        public static SelectionMgr SelectMgr => ActiveModel.SelectionManager;
        public static SelectionMgr SelectType => SelectMgr.GetSelectedObject6(1, -1);

        //public static object GetSelectionMgr()
        //{

        //    return null;
        //}
    }
}
