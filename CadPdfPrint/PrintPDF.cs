using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System.Threading;

namespace CadPdfPrint
{
    //为多线程用法1定义的类，暂时废弃
    class PrintPDF
    {
        private String PdfName;
        private Document acDoc;
        private Layout acLayout;
        private PlotInfo acPlInfo;
        private Editor acEd;

        public PrintPDF(String PdfName, Document acDoc, Layout acLayout, PlotInfo acPlInfo, Editor acEd)
        {
            this.PdfName = PdfName;
            this.acDoc = acDoc;
            this.acLayout = acLayout;
            this.acPlInfo = acPlInfo;
            this.acEd = acEd;
        }

        public void Printing()
        {
            while (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
            {
                Thread.Sleep(1000);
            }

            using (PlotEngine acPlEng = PlotFactory.CreatePublishEngine())
            {

                // 使用PlotProgressDialog对话框跟踪打印进度
                PlotProgressDialog acPlProgDlg = new PlotProgressDialog(false, 1, true);

                try
                {
                    using (acPlProgDlg)
                    {
                        // 定义打印开始时显示的状态信息
                        acPlProgDlg.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Plot Progress");
                        acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                        acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                        acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                        acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");

                        // 设置打印进度范围
                        acPlProgDlg.LowerPlotProgressRange = 0;
                        acPlProgDlg.UpperPlotProgressRange = 100;
                        acPlProgDlg.PlotProgressPos = 0;
                        // 显示打印进度对话框
                        acPlProgDlg.OnBeginPlot();
                        acPlProgDlg.IsVisible = true;
                        // 开始打印
                        acPlEng.BeginPlot(acPlProgDlg, null);
                        // 定义打印输出
                        //if (PdfName.Length > 0)
                        //{
                        //    PdfName = getPdfFileName(acDoc.Name,"", PdfName);

                        //}
                        //else
                        //{
                        PdfName = GetPdfFileName(acDoc.Name, acLayout.LayoutName, PdfName);
                        //acPlEng.BeginDocument(acPlInfo, acDoc.Name, null, 1, true, Path.GetDirectoryName(acDoc.Name) + "\\" + Path.GetFileNameWithoutExtension(acDoc.Name) + " - " + acLayout.LayoutName + ".pdf");
                        //}
                        acPlEng.BeginDocument(acPlInfo, acDoc.Name, null, 1, true, PdfName);
                        //acPlEng.BeginDocument(acPlInfo, acDoc.Name, null, 1, true, "c:\\temp\\1.pdf");
                        // 显示当前打印任务的有关信息
                        acPlProgDlg.set_PlotMsgString(PlotMessageIndex.Status, "Plotting: " + acDoc.Name + " - " + acLayout.LayoutName);
                        // 设置图纸进度范围
                        acPlProgDlg.OnBeginSheet();
                        acPlProgDlg.LowerSheetProgressRange = 0;
                        acPlProgDlg.UpperSheetProgressRange = 100;
                        acPlProgDlg.SheetProgressPos = 0;
                        // 打印第一张图/布局
                        PlotPageInfo acPlPageInfo = new PlotPageInfo();
                        acPlEng.BeginPage(acPlPageInfo, acPlInfo, true, null);
                        acPlEng.BeginGenerateGraphics(null);
                        acPlEng.EndGenerateGraphics(null);
                        // 结束第一张图/布局的打印
                        acPlEng.EndPage(null);
                        acPlProgDlg.SheetProgressPos = 100;
                        acPlProgDlg.OnEndSheet();
                        // 结束文档局的打印
                        acPlEng.EndDocument(null);
                        // 打印结束
                        acPlProgDlg.PlotProgressPos = 100;
                        acPlProgDlg.OnEndPlot();
                        //acPlEng.EndPlot(null);
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    acEd.WriteMessage("/n Exception message :" + ex.Message);
                }
                finally
                {
                    acPlEng.EndPlot(null);
                }
            }

        }

        public static String GetPdfFileName(String DocName, String LayoutName, String PdfName)
        {
            if (PdfName.Length > 0)
            {
                if (PdfName.Contains("/"))
                {
                    PdfName = PdfName.Replace('/', '-');
                }


                PdfName = Path.GetDirectoryName(DocName) + "\\" + PdfName + ".pdf";
            }
            else
            {
                PdfName = Path.GetDirectoryName(DocName) + "\\" + Path.GetFileNameWithoutExtension(DocName) + " - " + LayoutName + ".pdf";
            }

            if (File.Exists(PdfName))
            {
                PdfName = Path.GetDirectoryName(PdfName) + "\\" + Path.GetFileNameWithoutExtension(PdfName) + "(" + DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss") + ").pdf";
            }

            return PdfName;
        }
    }
}
