using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace CadPdfPrint
{
    public class PdfPrintClass
    {
        public const String RooneyTitle = "C  ROONEY EARTHMOVING LTD";
        public const String RooneyTitle1 = "C   ROONEY EARTHMOVING LTD";

        [CommandMethod("PdfPrintModel")]
        public static void PdfPrintModel()
        {
            // 获取当前文档和数据库，启动事务
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor acEd = acDoc.Editor;



            /*
            //1. 提示用户选择全部打印还是选择打印范围
            PromptKeywordOptions pKeyOpstsAllorSel = new PromptKeywordOptions("");
            pKeyOpstsAllorSel.Message = "\nPrint <All> or <Selected> Plans in Model Space to PDF?";
            pKeyOpstsAllorSel.Keywords.Add("All plans");
            pKeyOpstsAllorSel.Keywords.Add("Selected plans");
            pKeyOpstsAllorSel.Keywords.Default = "All plans";
            pKeyOpstsAllorSel.AllowNone = true;

            PromptResult pKeyRes = acDoc.Editor.GetKeywords(pKeyOpstsAllorSel);

            //如果用户按了Esc键，就退出
            if (pKeyRes.Status == PromptStatus.Cancel) return;

            if (pKeyRes.StringResult.Equals("Selected") || pKeyRes.StringResult.Equals("Selected plans"))
            {
                
                PromptPointOptions pPtOpts = new PromptPointOptions("");

                pPtOpts.Message = "\nEnter the start point of print area:";
                PromptPointResult pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                Point3d ptPrintAreaStart = pPtRes.Value;

                //如果用户按了Esc键或取消命令，就退出
                if (pPtRes.Status == PromptStatus.Cancel) return;

                pPtOpts.Message = "\nEnter the end point of print area:";
                //pPtOpts.UseDashedLine
                pPtOpts.UseBasePoint = true;
                pPtOpts.UseDashedLine = true;
                pPtOpts.BasePoint = ptPrintAreaStart;


                pPtRes = acEd.GetPoint(pPtOpts);
                Point3d ptPrintAreaEnd = pPtRes.Value;

                //如果用户按了Esc键或取消命令，就退出
                if (pPtRes.Status == PromptStatus.Cancel) return;

                Application.ShowAlertDialog("Print Area is:\n" + ptPrintAreaStart.X + " " + ptPrintAreaStart.Y + "\n to " + ptPrintAreaEnd.X + " " + ptPrintAreaEnd.Y);

            }
            */
            //Application.ShowAlertDialog("\nYou selected "+pKeyRes.StringResult);



            try
            {

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    // 引用布局管理器LayoutManager
                    LayoutManager acLayoutMgr;
                    acLayoutMgr = LayoutManager.Current;

                    //是否全部打印的标志，true：则是全部打印，flase：用户选择部分打印
                    bool isPrintAll = true;
                    

                    List<PrintAreaAndTitle> lstPrintAreaAndTitle = new List<PrintAreaAndTitle>();

                    //如果当前不是Model空间，则切换到Model空间
                    if (!acLayoutMgr.CurrentLayout.Equals("Model"))
                    {

                        //acEd.SwitchToModelSpace();
                        //Application.SetSystemVariable("CVPORT", 2);
                        //Application.ShowAlertDialog("Please run this command in Model Space, thanks.");

                        //LayoutManager.Current.CurrentLayout = "Model";
                        //Application.SetSystemVariable("TILEMODE", 1);
                        
                        acLayoutMgr.CurrentLayout = "Model";
                    }

                    // 读取当前布局，在命令行窗口显示布局名
                    Layout acLayout;
                    acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;
                                        
                    //如果用户已经选择了打印区域
                    PromptSelectionResult promptSelectionResult = acEd.SelectImplied();
                    SelectionSet sSet;

                    //如果提示状态OK，说明启动命令前选择了对象
                    if (promptSelectionResult.Status == PromptStatus.OK)
                    {
                        sSet = promptSelectionResult.Value;
                        if (sSet == null) return;

                        //部分打印
                        isPrintAll = false;
                        lstPrintAreaAndTitle = GetModelPrintAreaAndTileSelected(acCurDb, acEd,sSet);
                    }
                    else
                    {

                        //1. 提示用户选择全部打印还是选择打印范围
                        PromptKeywordOptions pKeyOpstsAllorSel = new PromptKeywordOptions("");
                        pKeyOpstsAllorSel.Message = "\nPrint <All> or <Selected> Plans in Model Space to PDF?";
                        pKeyOpstsAllorSel.Keywords.Add("All plans");
                        pKeyOpstsAllorSel.Keywords.Add("Selected plans");
                        pKeyOpstsAllorSel.Keywords.Default = "All plans";
                        pKeyOpstsAllorSel.AllowNone = true;

                        PromptResult pKeyRes = acDoc.Editor.GetKeywords(pKeyOpstsAllorSel);

                        //如果用户按了Esc键，就退出
                        if (pKeyRes.Status == PromptStatus.Cancel) return;

                        if (pKeyRes.StringResult.Equals("Selected") || pKeyRes.StringResult.Equals("Selected plans"))
                        {
                            //请求从图形区域选择对象：
                            acEd.TurnSubentityWindowSelectionOn();
                            promptSelectionResult = acEd.GetSelection();

                            //如果提示状态OK，表示已经选择好对象
                            if (promptSelectionResult.Status == PromptStatus.OK)
                            {
                                sSet = promptSelectionResult.Value;

                                if (sSet == null) return;

                                //部分打印
                                isPrintAll = false;
                                lstPrintAreaAndTitle = GetModelPrintAreaAndTileSelected(acCurDb, acEd, sSet);
                            }

                            //如果用户按了Esc，就退出
                            if (promptSelectionResult.Status == PromptStatus.Cancel)
                            {
                                return;
                            }
                        }
                        
                    }
                    

                    if(isPrintAll == true)
                    {
                        //如果是打印全部
                        lstPrintAreaAndTitle = GetModelPrintAreaAndTitle(acCurDb, acEd);
                    }
                    

                    if (lstPrintAreaAndTitle.Count < 1)
                    {
                        Application.ShowAlertDialog("\nThere is nothing to print.");
                        return;
                    }

                    foreach(PrintAreaAndTitle paat in lstPrintAreaAndTitle)
                    {
                        // 输出当前布局名和设备名
                        acDoc.Editor.WriteMessage("\nCurrent layout: " + acLayout.LayoutName);
                        acDoc.Editor.WriteMessage("\nCurrent device name: " + acLayout.PlotConfigurationName);
                        // 从布局中获取PlotInfo
                        PlotInfo acPlInfo = new PlotInfo();
                        acPlInfo.Layout = acLayout.ObjectId;

                            // 复制布局中的PlotSettings
                            PlotSettings acPlSet = new PlotSettings(acLayout.ModelType);
                            acPlSet.CopyFrom(acLayout);

                            //更新PlotSetting对象
                            PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                        //设置打印区域

                        //PromptPointOptions pPtOpts = new PromptPointOptions("");

                        //pPtOpts.Message = "\nEnter the start point of print area:";
                        //PromptPointResult pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                        //Autodesk.AutoCAD.Geometry.Point3d ptPrintAreaStart = pPtRes.Value;

                        ////如果用户按了Esc键或取消命令，就退出
                        //if (pPtRes.Status == PromptStatus.Cancel) return;

                        //pPtOpts.Message = "\nEnter the end point of print area:";
                        ////pPtOpts.UseDashedLine
                        //pPtOpts.UseBasePoint = true;
                        //pPtOpts.UseDashedLine = true;
                        //pPtOpts.BasePoint = ptPrintAreaStart;


                        //pPtRes = acEd.GetPoint(pPtOpts);
                        //Point3d ptPrintAreaEnd = pPtRes.Value;

                        ////如果用户按了Esc键或取消命令，就退出
                        //if (pPtRes.Status == PromptStatus.Cancel) return;
                        try
                        {
                            //转换WCS的点坐标到DCS的点坐标
                            //Extents2d window = TransWCS2DCS(acDoc, lstPrintAreaAndTitle[0].ptPrintAreaLeftUp, lstPrintAreaAndTitle[0].ptPrintAreaRightDown);
                            Extents2d window = TransWCS2DCS(acDoc, paat.ptPrintAreaLeftUp, paat.ptPrintAreaRightDown);



                            //获取并转换Layout名 如 192/2/1  =>  192-2-1
                            //String PdfName = getLayoutName(acCurDb, acEd);

                            //设置打印区域坐标
                            acPlSetVdr.SetPlotWindowArea(acPlSet, window);
                            //设置打印区域为Window打印
                            acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);


                            acPlSetVdr.SetUseStandardScale(acPlSet, true);
                            acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);

                            acPlSetVdr.SetPlotCentered(acPlSet, true);


                            //读取当前打印机信息
                            StringCollection devlist = acPlSetVdr.GetPlotDeviceList();


                            String WebPdfPrinter = "Web";
                            bool IsFoundWebPdfPrinter = false;

                            for (int i = 0; i < devlist.Count; i++)
                            {
                                if (devlist[i].Contains(WebPdfPrinter))
                                {
                                    WebPdfPrinter = devlist[i];
                                    IsFoundWebPdfPrinter = true;
                                    break;
                                }
                            }

                            if (IsFoundWebPdfPrinter != true)
                            {
                                String PdfPrinter = "Pdf";
                                for (int i = 0; i < devlist.Count; i++)
                                {
                                    if (devlist[i].Contains(PdfPrinter))
                                    {
                                        WebPdfPrinter = devlist[i];
                                        break;
                                    }
                                }

                            }

                            //
                            var plotConfig = PlotConfigManager.SetCurrentConfig(WebPdfPrinter);
                            
                            // 【1.用Papersize = null 等参数初始化使用的打印设备】
                            try
                            {
                                acPlSetVdr.SetPlotConfigurationName(acPlSet, WebPdfPrinter, null);
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                acEd.WriteMessage("/n Exception message :" + ex.Message);
                                acPlSet.Dispose();
                                return;
                            }

                            //【2. 设备初始化好以后，再获取可用纸张size】
                            StringCollection CanonicalMediaNameList = acPlSetVdr.GetCanonicalMediaNameList(acPlSet);
                            String PdfPaperSize = "A3";

                            for (int i = 0; i < CanonicalMediaNameList.Count; i++)
                            {
                                if (CanonicalMediaNameList[i].Contains(PdfPaperSize))
                                {
                                    PdfPaperSize = CanonicalMediaNameList[i];
                                    continue;
                                }
                            }
                            

                            // 【3.再用2获取到的纸张size设置打印设备。至此，打印机类型和纸张才设定完毕。】
                            try
                            {
                                acPlSetVdr.SetPlotConfigurationName(acPlSet, WebPdfPrinter, PdfPaperSize);
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                Application.ShowAlertDialog("/n Exception message :" + ex.Message);
                                //acEd.WriteMessage("/n Exception message :" + ex.Message);
                                //acPlSet.Dispose();
                            }



                            //设置打印方向；
                            var mediaBounds = plotConfig.GetMediaBounds(PdfPaperSize);
                            var layoutWidth = acLayout.Extents.MaxPoint.X - acLayout.Extents.MinPoint.X;
                            var layoutHeight = acLayout.Extents.MaxPoint.Y - acLayout.Extents.MinPoint.Y;

                            PlotRotation plotRotation = PlotRotation.Degrees000;
                            if (layoutWidth > layoutHeight && mediaBounds.UpperRightPrintableArea.X < mediaBounds.UpperRightPrintableArea.Y)
                                plotRotation = PlotRotation.Degrees090;
                            else if (layoutWidth < layoutHeight && mediaBounds.UpperRightPrintableArea.X > mediaBounds.UpperRightPrintableArea.Y)
                                plotRotation = PlotRotation.Degrees090;

                            acPlSetVdr.SetPlotRotation(acPlSet, plotRotation);
                            //acPlSetVdr.SetPlotRotation(acPlSet, GetPlotRotation(plotConfig, PdfPaperSize, acLayout));


                            // 用上述设置信息覆盖PlotInfo对象，
                            // 不会将修改保存回布局
                            acPlInfo.OverrideSettings = acPlSet;
                            // 验证打印信息
                            PlotInfoValidator acPlInfoVdr = new PlotInfoValidator();
                            acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;

                            acPlInfoVdr.Validate(acPlInfo);

                            while (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
                            {
                                Thread.Sleep(1000);
                            }

                            //保存App的原参数
                            short bgPlot = (short)Application.GetSystemVariable("BACKGROUNDPLOT");
                            //设定为前台打印，加快打印速度
                            Application.SetSystemVariable("BACKGROUNDPLOT", 0);
                            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                            {
                                using (PlotEngine acPlEng = PlotFactory.CreatePublishEngine())
                                {
                                    // 使用PlotProgressDialog对话框跟踪打印进度
                                    PlotProgressDialog acPlProgDlg = new PlotProgressDialog(false, 1, true);
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
                                        String PdfName = GetPdfFileName(acDoc.Name, acLayout.LayoutName, paat.PdfTitle);

                                        acPlEng.BeginDocument(acPlInfo, acDoc.Name, null, 1, true, PdfName);

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

                                        //打印完毕，把系统参数复原
                                        Application.SetSystemVariable("BACKGROUNDPLOT", bgPlot);
                                        acPlEng.EndPlot(null);
                                        acPlSet.Dispose();

                                    }
                                }
                            }
                        }
                        catch(Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            
                            acPlSet.Dispose();
                            acEd.WriteMessage("/n Exception message :", ex.Message);
                        }
                    }
                    acTrans.Commit();
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                acEd.WriteMessage("/n Exception message :", ex.Message);
            }
            //finally
            //{
            //    //打印完毕，把系统参数复原
            //    Application.SetSystemVariable("BACKGROUNDPLOT", bgPlot);

            //    acPlSet.Dispose();
            //    acPlEng.EndPlot(null);
            //}
        }

        
        [CommandMethod("PdfPrint")]
        public static void PdfPrint()
        { 
        //[LispFunction(("PdfPrint"))]
        //public static void PdfPrint(ResultBuffer rbArgs)
        //{
        //    //取得参数

        //    string strVal = "";

        //    if (rbArgs != null)
        //    {
                
        //        int nCnt = 0;
        //        foreach (TypedValue rb in rbArgs)
        //        {
        //            //if (rb.TypeCode == (int)Autodesk.AutoCAD.Runtime.LispDataType.Text && nCnt.Equals(0) && rb.Value.ToString.Equals("p"))
        //            if (rb.TypeCode == (int)Autodesk.AutoCAD.Runtime.LispDataType.Text)
        //            {
        //                strVal = rb.Value.ToString();
        //                nCnt = nCnt + 1;
        //            }
        //        }

        //    }

            ////默认设置打印方向为横向，如果参数为“P”或者“p”，那么设置为纵向打印
            //PlotRotation PrintRotataion = PlotRotation.Degrees000;
            //if (strVal.ToUpper().Equals("P"))
            //{
            //    PrintRotataion = PlotRotation.Degrees090;
            //}


            // 获取当前文档和数据库，启动事务
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor acEd = acDoc.Editor;

            //获取图纸编号 如 192/2/1 
            String PdfName = getLayoutName(acCurDb, acEd);

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // 引用布局管理器LayoutManager
                LayoutManager acLayoutMgr;
                acLayoutMgr = LayoutManager.Current;
                // 获取当前布局，在命令行窗口显示布局名字
                Layout acLayout;
                acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;

                if(acLayout.LayoutName.Equals("Model"))
                {
                    Application.ShowAlertDialog("Cannot run this command in Model Space, please switch to Layout Space.");
                    return;
                }

                PlotRotation printRotation = PlotRotation.Degrees000;

                //提示用户选择横向打印Landscape，还是纵向打印Portrait
                PromptKeywordOptions pKeyOpsts = new PromptKeywordOptions("");
                pKeyOpsts.Message = "\nPlease select drawing orientation :";
                pKeyOpsts.Keywords.Add("Landscape");
                pKeyOpsts.Keywords.Add("Portrait");
                pKeyOpsts.Keywords.Default = "Landscape";
                pKeyOpsts.AllowNone = true;

                PromptResult pKeyRes = acDoc.Editor.GetKeywords(pKeyOpsts);

                //如果用户按了Esc键，就退出
                if (pKeyRes.Status == PromptStatus.Cancel) return;

                if (pKeyRes.StringResult.Equals("Portrait"))
                {
                    printRotation = PlotRotation.Degrees090;
                }

                // 从布局中获取PlotInfo
                PlotInfo acPlInfo = new PlotInfo();
                acPlInfo.Layout = acLayout.ObjectId;
                // 复制布局中的PlotSettings
                PlotSettings acPlSet = new PlotSettings(acLayout.ModelType);
                acPlSet.CopyFrom(acLayout);

                // 更新PlotSettings对象
                PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                //读取当前打印机信息
                StringCollection devlist = acPlSetVdr.GetPlotDeviceList();

                String WebPdfPrinter = "Web";
                bool IsFoundWebPdfPrinter = false;
                for (int i = 0; i < devlist.Count; i++)
                {
                    if (devlist[i].Contains(WebPdfPrinter))
                    {
                        WebPdfPrinter = devlist[i];
                        IsFoundWebPdfPrinter = true;
                        break;
                    }
                }

                if(IsFoundWebPdfPrinter!=true)
                {
                    String PdfPrinter = "Pdf";
                    for (int i = 0; i < devlist.Count; i++)
                    {
                        if (devlist[i].Contains(PdfPrinter))
                        {
                            WebPdfPrinter = devlist[i];
                            break;
                        }
                    }

                }
                

                //
                var plotConfig = PlotConfigManager.SetCurrentConfig(WebPdfPrinter);

                // 【1.用Papersize = null 等参数初始化使用的打印设备】
                try
                {
                    acPlSetVdr.SetPlotConfigurationName(acPlSet, WebPdfPrinter, null);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    acEd.WriteMessage("/n Exception message :" + ex.Message);
                    acPlSet.Dispose();
                    return;
                }

                //【2. 设备初始化好以后，再获取可用纸张size】
                StringCollection CanonicalMediaNameList = acPlSetVdr.GetCanonicalMediaNameList(acPlSet);
                String PdfPaperSize = "A3";

                for (int i = 0; i < CanonicalMediaNameList.Count; i++)
                {
                    if (CanonicalMediaNameList[i].Contains(PdfPaperSize))
                    {
                        PdfPaperSize = CanonicalMediaNameList[i];
                        continue;
                    }
                }

                // 【3.再用2获取到的纸张size设置打印设备。至此，打印机类型和纸张才设定完毕。】
                try
                {
                    acPlSetVdr.SetPlotConfigurationName(acPlSet, WebPdfPrinter, PdfPaperSize);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    acEd.WriteMessage("/n Exception message :" + ex.Message);
                    //acPlSet.Dispose();
                }



                //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

                // 设置打印区域（即打印对话框中的Plot Area）
                //注意：如果打印区域设为Layout，那么下面就不能设置“居中”，因为二者是互斥的。
                //acPlSetVdr.SetPlotType(acPlSet,Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);

                // 设置打印比例
                //acPlSetVdr.SetUseStandardScale(acPlSet, true);
                //acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);

                // 居中打印
                //  One cannot Plot a Layout Area Centered, i.e. the combination of Layout and Centered is not allowed.
                //acPlSetVdr.SetPlotCentered(acPlSet, true);

                //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
                //设置旋转方向
                //acPlSetVdr.SetPlotRotation(acPlSet, GetPlotRotation(plotConfig,PdfPaperSize,acLayout));

                //设置打印方向，默认为Landscape/横向，如果命令参数为P或者p，那么打印方向设置为Portrait/纵向
                acPlSetVdr.SetPlotRotation(acPlSet, printRotation);

                // 用上述设置信息覆盖PlotInfo对象，
                // 不会将修改保存回布局
                acPlInfo.OverrideSettings = acPlSet;
                // 验证打印信息
                PlotInfoValidator acPlInfoVdr = new PlotInfoValidator();
                acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                try
                {
                    acPlInfoVdr.Validate(acPlInfo);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    acPlSet.Dispose();
                    acEd.WriteMessage("/n Exception message :" + ex.Message);
                }

                // 检查是否有正在处理的打印任务

                //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/  多线程打印，暂时废弃 -Start

                /////传参数多线程用法1，新定义一个类，把参数定义为该类的成员变量，在构造函数中传递参数并初始化；另外定义主线程函数，可以在其间操作该类的各种成员函数，从而避免直接传递参数
                //PrintPDF printPDF = new PrintPDF(PdfName, acDoc, acLayout, acPlInfo, acEd);
                //Thread th = new Thread(printPDF.Printing);
                //th.Start();

                //////传参数多线程用法2
                //Thread th1 = new Thread(() =>  Printing(PdfName, acDoc, acLayout, acPlInfo, acEd));
                //th1.Start();

                //_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/  多线程打印，暂时废弃 -End

                while (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
                {
                    Thread.Sleep(1000);
                }

                //保存App的原参数
                short bgPlot =(short)Application.GetSystemVariable("BACKGROUNDPLOT");
                //设定为前台打印，加快打印速度
                Application.SetSystemVariable("BACKGROUNDPLOT", 0);
                    if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                    {
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

                                    //用“图纸编号”、“图纸文件名及路径”等信息，对生成的PDF文件名进行格式化
                                    PdfName = GetPdfFileName(acDoc.Name, acLayout.LayoutName, PdfName);

                                    acPlEng.BeginDocument(acPlInfo, acDoc.Name, null, 1, true, PdfName);
                                    
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
                            //打印完毕，把系统参数复原
                            Application.SetSystemVariable("BACKGROUNDPLOT",bgPlot);

                            acPlSet.Dispose();
                            acPlEng.EndPlot(null);
                            }
                        }

                     
                    }
                    else
                    {
                        Thread.Sleep(1000);
                    }
                                
            }
        }

        public static Extents2d TransWCS2DCS(Document acDoc, Point3d ptPrintAreaStart, Point3d ptPrintAreaEnd)
        {
            // Gets the current view
            ViewTableRecord acView = acDoc.Editor.GetCurrentView();

            //1.因为Extents3d初始化时，要求第一个坐标小于第二个坐标，所以先进行一下判断&交换坐标值
            //2.交换好坐标之后，为了能把图纸边界线也打印出来，需要把打印区域扩大一点，具体做法就是把左下角点的坐标减去一个系数，右上角点的坐标加上一个系数
            double factor = Math.Abs(ptPrintAreaStart.X - ptPrintAreaEnd.X) / 300;

            double smallX = Math.Min(ptPrintAreaStart.X,ptPrintAreaEnd.X) - factor;
            double bigX = Math.Max(ptPrintAreaStart.X, ptPrintAreaEnd.X) + factor;
            double smallY = Math.Min(ptPrintAreaStart.Y, ptPrintAreaEnd.Y) - factor;
            double bigY = Math.Max(ptPrintAreaStart.Y, ptPrintAreaEnd.Y)+ factor;
            double smallZ = Math.Min(ptPrintAreaStart.Z, ptPrintAreaEnd.Z);
            double bigZ = Math.Max(ptPrintAreaStart.Z, ptPrintAreaEnd.Z);

            //Point3d ptStart


            Extents3d eExtents = new Extents3d(new Point3d(smallX,smallY,smallZ), new Point3d(bigX,bigY,bigZ));

            //Extents3d eExtents = new Extents3d();
            //eExtents.AddPoint(ptPrintAreaStart);
            //eExtents.AddPoint(ptPrintAreaEnd);

            //eExtents.MinPoint.Add(new Vector3d(factor, factor, 0));
            //eExtents.MaxPoint.Add(new Vector3d(factor, factor, 0));


            //Extents3d eExtents = new Extents3d(new Point3d(1472513.91797399, 5084804.00380163, ptPrintAreaStart.Z), new Point3d(1490499.70879218, 5097282.14443037, ptPrintAreaEnd.Z));
            // Translates WCS coordinates to DCS
            Matrix3d matWCS2DCS;
            matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
            matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS;
            matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                                           acView.ViewDirection,
                                           acView.Target) * matWCS2DCS;

            // Tranforms the extents to DCS
            matWCS2DCS = matWCS2DCS.Inverse();
            eExtents.TransformBy(matWCS2DCS);

            return new Extents2d(eExtents.MinPoint.X, eExtents.MinPoint.Y, eExtents.MaxPoint.X, eExtents.MaxPoint.Y);
        }

        //Plot All PaperSpace Layouts of All Opened Documents to PDF and
        // Optional Change Plot Color: Color.ctb
        // Optional Change Plot Size: A3 and implicit Plot Best Fit
        [CommandMethod("PAL", CommandFlags.Session)]
        public static void PlotAllLayouts()
        {
            Document doc = null;
            Database db = null;
            Editor ed = null;


            //Required: Plot to the Device
            string plotDeviceName = "DWG To PDF.pc3";

            //Optional: specify NULL or Change output Colors
            string ctbFilename = "Color.ctb";

            //Optional specify NULL or Change plot Size, and implicit Plot ToBestFit
            string plotPaperName = "A3";


            try
            {
                dynamic acadApp = Application.AcadApplication;

                doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                    throw new System.Exception("No MdiActiveDocument");
                db = doc.Database;
                ed = doc.Editor;

                LayoutManager lm = LayoutManager.Current;
                PlotSettingsValidator psv = PlotSettingsValidator.Current;
                PlotConfigManager.RefreshList(RefreshCode.All);

                if (!string.IsNullOrWhiteSpace(ctbFilename))
                {
                    var plotStyles = psv.GetPlotStyleSheetList().Cast<string>().ToList();
                    if (!plotStyles.Contains(ctbFilename, StringComparer.OrdinalIgnoreCase))
                        throw new System.Exception("CTB not found: " + ctbFilename);
                }

                var plotDevices = psv.GetPlotDeviceList().Cast<string>().ToList();
                if (!plotDevices.Contains(plotDeviceName, StringComparer.OrdinalIgnoreCase))
                    throw new System.Exception("PC3 not found: " + plotDeviceName);
                var plotConfig = PlotConfigManager.SetCurrentConfig(plotDeviceName);

                plotConfig.RefreshMediaNameList();
                string canonicalMediaName = null;
                foreach (string mediaName in plotConfig.CanonicalMediaNames)
                {
                    string localMediaName = plotConfig.GetLocalMediaName(mediaName);
                    if (localMediaName == plotPaperName)
                    {
                        canonicalMediaName = mediaName;
                        break;
                    }
                }
                if (string.IsNullOrWhiteSpace(canonicalMediaName))
                    throw new System.Exception("Plot Paper not found: " + plotPaperName);


                foreach (Document openDoc in Application.DocumentManager)
                {
                    if (!openDoc.IsActive)
                        Application.DocumentManager.MdiActiveDocument = openDoc;

                    using (var openDocLock = openDoc.LockDocument())
                    {
                        Database openDb = openDoc.Database;
                        Editor openEd = openDoc.Editor;
                        string openDocName = Path.GetFileNameWithoutExtension(openDoc.Name);
                        openEd.WriteMessage("\n {0}", openDoc.Name);
                        string openDocPath = Path.GetTempPath();
                        if (openDoc.IsNamedDrawing)
                            openDocPath = Path.GetDirectoryName(openDoc.Name);

                        using (Transaction tr = openDb.TransactionManager.StartTransaction())
                        {
                            DBDictionary layoutDict = (DBDictionary)tr.GetObject(openDb.LayoutDictionaryId, OpenMode.ForRead);
                            foreach (DBDictionaryEntry layoutEntry in layoutDict)
                            {
                                if (layoutEntry.Key.ToUpper() == "MODEL")
                                    continue;


                                string layoutName = layoutEntry.Key;
                                ObjectId layoutId = layoutEntry.Value;
                                lm.CurrentLayout = layoutName;
                                acadApp.ZoomExtents();
                                openEd.WriteMessage("\n\t {0}", layoutName);


                                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                                bool plotToFile = false;
                                string pdfPathname = null;

                                //Change Plot Settings, i.e. PlotDevice, PlotStyle, PaperSize
                                using (PlotSettings tmpPlotSettings = new PlotSettings(layout.ModelType))
                                {
                                    tmpPlotSettings.CopyFrom(layout);
                                    
                                    //Do we need to change the PlotDevice?
                                    if (!string.IsNullOrWhiteSpace(plotDeviceName) && tmpPlotSettings.PlotConfigurationName != plotDeviceName)
                                        psv.SetPlotConfigurationName(tmpPlotSettings, plotDeviceName, null);

                                    if (Regex.IsMatch(tmpPlotSettings.PlotConfigurationName, "PDF", RegexOptions.IgnoreCase))
                                    {
                                        plotToFile = true;
                                        pdfPathname = Path.Combine(openDocPath, openDocName + "-" + layoutName + ".pdf");
                                        if (File.Exists(pdfPathname))
                                            File.Delete(pdfPathname);
                                    }

                                    //Do we need to change the PlotColor?
                                    if (!string.IsNullOrWhiteSpace(ctbFilename) && tmpPlotSettings.CurrentStyleSheet != ctbFilename)
                                    {
                                        psv.SetCurrentStyleSheet(tmpPlotSettings, ctbFilename);
                                        tmpPlotSettings.PlotPlotStyles = true;
                                        tmpPlotSettings.ShowPlotStyles = true;
                                    }

                                    //Do we need to change the PlotSize, then Rotate for Best Fit, set ScaleToFit and set centered
                                    if (!string.IsNullOrWhiteSpace(plotPaperName))
                                    {
                                        psv.SetCanonicalMediaName(tmpPlotSettings, canonicalMediaName);
                                        var mediaBounds = plotConfig.GetMediaBounds(canonicalMediaName);
                                        var layoutWidth = layout.Extents.MaxPoint.X - layout.Extents.MinPoint.X;
                                        var layoutHeight = layout.Extents.MaxPoint.Y - layout.Extents.MinPoint.Y;

                                        PlotRotation plotRotation = PlotRotation.Degrees000;
                                        if (layoutWidth > layoutHeight && mediaBounds.UpperRightPrintableArea.X < mediaBounds.UpperRightPrintableArea.Y)
                                            plotRotation = PlotRotation.Degrees090;
                                        else if (layoutWidth < layoutHeight && mediaBounds.UpperRightPrintableArea.X > mediaBounds.UpperRightPrintableArea.Y)
                                            plotRotation = PlotRotation.Degrees090;
                                        psv.SetPlotRotation(tmpPlotSettings, plotRotation);

                                        psv.SetPlotType(tmpPlotSettings, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                                        psv.SetUseStandardScale(tmpPlotSettings, true);
                                        psv.SetStdScaleType(tmpPlotSettings, StdScaleType.ScaleToFit);
                                        psv.SetPlotCentered(tmpPlotSettings, true);
                                        psv.RefreshLists(tmpPlotSettings);
                                    }

                                    layout.UpgradeOpen();
                                    layout.CopyFrom(tmpPlotSettings);
                                }


                                //Finally Plot the Layout
                                using (var plotInfo = new PlotInfo())
                                {
                                    using (PlotEngine plotEngine = PlotFactory.CreatePublishEngine())
                                    {
                                        plotEngine.BeginPlot(null, null);
                                        PlotInfoValidator validator = new PlotInfoValidator();
                                        validator.MediaMatchingPolicy = Autodesk.AutoCAD.PlottingServices.MatchingPolicy.MatchEnabled;
                                        plotInfo.Layout = layoutId;
                                        validator.Validate(plotInfo);
                                        plotEngine.BeginDocument(plotInfo, openDoc.Name, null, 1, plotToFile, pdfPathname);
                                        PlotPageInfo pageInfo = new PlotPageInfo();
                                        plotEngine.BeginPage(pageInfo, plotInfo, true, null);
                                        plotEngine.BeginGenerateGraphics(null);
                                        plotEngine.EndGenerateGraphics(null);
                                        plotEngine.EndPage(null);
                                        plotEngine.EndDocument(null);
                                        plotEngine.EndPlot(null);
                                    }
                                }
                                if (plotToFile && File.Exists(pdfPathname))
                                    openEd.WriteMessage("\n\t\t {0}", pdfPathname);

                            }
                            tr.Commit();
                        }
                    }
                }

                //return to the previous active document
                if (!doc.IsActive)
                    Application.DocumentManager.MdiActiveDocument = doc;

            }
            catch (System.Exception ex)
            {
                if (ed != null)
                {
                    if (!doc.IsActive)
                        Application.DocumentManager.MdiActiveDocument = doc;
                    ed.WriteMessage("\n Error in PlotAllLayouts: {0}", ex.Message);
                }
                else
                    Application.ShowAlertDialog(string.Format("Error in PlotAllLayouts: {0}", ex.Message));
            }
        }


        public static PlotRotation GetPlotRotation(PlotConfig plotConfig,String PdfPaperSize,Layout acLayout)
        {
            var mediaBounds = plotConfig.GetMediaBounds(PdfPaperSize);
            var layoutWidth = acLayout.Extents.MaxPoint.X - acLayout.Extents.MinPoint.X;
            var layoutHeight = acLayout.Extents.MaxPoint.Y - acLayout.Extents.MinPoint.Y;

            PlotRotation plotRotation = PlotRotation.Degrees000;
            if (layoutWidth > layoutHeight && mediaBounds.UpperRightPrintableArea.X < mediaBounds.UpperRightPrintableArea.Y)
                plotRotation = PlotRotation.Degrees090;
            else if (layoutWidth < layoutHeight && mediaBounds.UpperRightPrintableArea.X > mediaBounds.UpperRightPrintableArea.Y)
                plotRotation = PlotRotation.Degrees090;

            return plotRotation;
        }

        public static List<PrintAreaAndTitle> GetModelPrintAreaAndTileSelected(Database acCurDb,Editor acEd, SelectionSet sSet)
        {
            List<PrintAreaAndTitle> lst = new List<PrintAreaAndTitle>();

            //此处定义的思路见GetModelPrintAreaAndTitle的说明
            //String RooneyTitle = "C  ROONEY EARTHMOVING LTD";
            int XTimes = 75;
            int YTimes = 15;

            List<DBText> lstTextPaperTitles = new List<DBText>();
            List<MText> lstMTextPaperTitles = new List<MText>();

            //String strLayoutName = "";
            String Rex1 = "^[0-9]+(/[0-9]+)+$";
            String Rex2 = "^[0-9]+(-[0-9]+)+$";


            //遍历选择集
            using(Transaction tr = acCurDb.TransactionManager.StartTransaction())
            {
                foreach(SelectedObject acSSObj in sSet)
                {
                    //以读模式打开所选对象
                    Entity acEnt  = tr.GetObject(acSSObj.ObjectId,OpenMode.ForRead) as Entity;
                    switch (acEnt.ObjectId.ObjectClass.DxfName)
                    {
                        case "TEXT":
                            DBText text = (DBText)acEnt;
                            if (Regex.IsMatch(text.TextString, Rex1) || Regex.IsMatch(text.TextString, Rex2))
                            {
                                lstTextPaperTitles.Add(text);
                            }
                            break;

                        case "MTEXT":
                            MText mtext = (MText)acEnt;
                            if (Regex.IsMatch(mtext.Text, Rex1) || Regex.IsMatch(mtext.Text, Rex2))
                            {
                                lstMTextPaperTitles.Add(mtext);
                            }
                            break;

                        default:
                            break;
                    }
                }
                var textClass = RXObject.GetClass(typeof(DBText));
                var mtextClass = RXObject.GetClass(typeof(MText));

                foreach (SelectedObject acSSObj in sSet)
                {
                    Entity acEnt = tr.GetObject(acSSObj.ObjectId,OpenMode.ForRead) as Entity;
                    if (acEnt.ObjectId.ObjectClass.IsDerivedFrom(textClass))
                    {
                        DBText text = tr.GetObject(acEnt.ObjectId,OpenMode.ForRead) as DBText;
                        if (text.TextString.Equals(RooneyTitle) || text.TextString.Equals(RooneyTitle1))
                        {
                            for (int i = 0; i < lstTextPaperTitles.Count;i++)
                            {
                                //如果该编号在公司名的合理偏移范围内，则认为该编号是正确的图纸编号
                                if ((Math.Abs(lstTextPaperTitles[i].Position.X - text.Position.X) < text.Height * XTimes) && (Math.Abs(lstTextPaperTitles[i].Position.Y - text.Position.Y) < text.Height * YTimes))
                                {
                                    //取得离公司名Text最近的Line，也就是图纸的【下边界线】
                                    Line line = GetNearestLine(acEnt.ObjectId);
                                    Point3d ptLeftUp = GetLeftUpPointFromLine(acCurDb, line);

                                    lst.Add(new PrintAreaAndTitle(lstTextPaperTitles[i].TextString, ptLeftUp, line.EndPoint ));
                                }
                            }
                        }
                    }
                    else if (acEnt.ObjectId.ObjectClass.IsDerivedFrom(mtextClass))
                    {
                        MText mtext = (MText)tr.GetObject(acEnt.ObjectId,OpenMode.ForRead);
                            if (mtext.Text.Equals(RooneyTitle) || mtext.Text.Equals(RooneyTitle1))
                            {
                                foreach (MText title in lstMTextPaperTitles)
                                {
                                    if((Math.Abs(title.Location.X - mtext.Location.X) < mtext.ActualHeight * XTimes) && (Math.Abs(title.Location.Y - mtext.Location.Y) < mtext.Height * YTimes))
                                    {
                                        Line line = GetNearestLine(acEnt.ObjectId);
                                        Point3d ptLeftUp = GetLeftUpPointFromLine(acCurDb, line);

                                        lst.Add(new PrintAreaAndTitle(title.Text, ptLeftUp, line.EndPoint));
                                        break;
                                    }
                                }
                            }
                    }
                }
                
            }

            return lst;
        }

        ////////////////////////////////////////
        //获取Model空间每个打印区域的图纸编号，及其范围坐标，左上点和右下点
        //图纸编号类似“193/2/2/1”或“187-123”
        //
        public static List<PrintAreaAndTitle> GetModelPrintAreaAndTitle(Database acCurDb, Editor acEd) //where PrintAreaAndTitle:new()
        {
            List<PrintAreaAndTitle> lst = new List<PrintAreaAndTitle>();

            //String RooneyTitle = "C  ROONEY EARTHMOVING LTD";

            //确定图纸编号的思路是：先找到公司名的TEXT，然后把公司名的高度作为标准，偏移一个区域，在该区域内再搜索是否存在表示图纸编号的TEXT，以此确保图纸编号的正确性
            //所以这里暂定两个整数，公司名高度分别乘以这两个整数，作为划定偏移区域的X和Y
            int XTimes = 75;
            int YTimes = 15;

            List<DBText> lstTextPaperTitles = new List<DBText>();
            List<MText> lstMTextPaperTitles = new List<MText>();

            //String strLayoutName = "";
            String Rex1 = "^[0-9]+(/[0-9]+)+$";
            String Rex2 = "^[0-9]+(-[0-9]+)+$";
            

            using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(acCurDb), OpenMode.ForRead);

                //第一遍循环，取出所有可能是图纸编号的TEXT/MTEXT，存下来。这样比每次都循环一遍和公司名来对比性能更好
                foreach(ObjectId id in modelSpace)
                {
                    switch (id.ObjectClass.DxfName)
                    {
                        case "TEXT":
                            DBText text = (DBText)tr.GetObject(id, OpenMode.ForRead);
                            if(Regex.IsMatch(text.TextString,Rex1) || Regex.IsMatch(text.TextString, Rex2))
                            {
                                lstTextPaperTitles.Add(text);
                            }
                            break;

                        case "MTEXT":
                            MText mtext = (MText)tr.GetObject(id, OpenMode.ForRead);
                            if (Regex.IsMatch(mtext.Text, Rex1)|| Regex.IsMatch(mtext.Text, Rex2))
                            {
                                lstMTextPaperTitles.Add(mtext);
                            }
                            break;

                        default:
                            break;

                    }
                }

                foreach (ObjectId id in modelSpace)
                {
                    //Tip：另外一个判断是否是Text的方法是  id.ObjectClass.IsDerivedFrom(textClass)
                    switch (id.ObjectClass.DxfName)
                    {
                        case "TEXT":
                            DBText text = (DBText)tr.GetObject(id, OpenMode.ForRead);
                            if (text.TextString.Equals( RooneyTitle) || text.TextString.Equals(RooneyTitle1))
                            {
                                for (int i=0; i < lstTextPaperTitles.Count; i++)
                                {
                                    //如果该编号在公司名的合理偏移范围内，则认为该编号是正确编号。否则连该公司名也Skip过去
                                    if ((Math.Abs(lstTextPaperTitles[i].Position.X - text.Position.X) < text.Height * XTimes) && (Math.Abs(lstTextPaperTitles[i].Position.Y - text.Position.Y) < text.Height * YTimes))
                                    {
                                        //取得离公司名Text最近的Line，也就是图纸的【下边界线】
                                        Line line = GetNearestLine(id);
                                        Point3d ptLeftUp = GetLeftUpPointFromLine(acCurDb, line);
                                        //PrintAreaAndTitle printAreaAndTitle = new PrintAreaAndTitle();
                                        
                                        lst.Add(new PrintAreaAndTitle(lstTextPaperTitles[i].TextString, ptLeftUp, line.EndPoint ));
                                        //break;
                                    }
                                }
                            }
                            break;
                        case "MTEXT":
                            MText mtext = (MText)tr.GetObject(id, OpenMode.ForRead);
                            if (mtext.Text.Equals(RooneyTitle) || mtext.Text.Equals(RooneyTitle1))
                            {
                                foreach (MText title in lstMTextPaperTitles)
                                {
                                    if((Math.Abs(title.Location.X - mtext.Location.X) < mtext.ActualHeight * XTimes) && (Math.Abs(title.Location.Y - mtext.Location.Y) < mtext.Height * YTimes))
                                    {
                                        Line line = GetNearestLine(id);
                                        Point3d ptLeftUp = GetLeftUpPointFromLine(acCurDb, line);

                                        lst.Add(new PrintAreaAndTitle(title.Text, ptLeftUp, line.EndPoint));
                                        break;
                                    }
                                }
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            return lst;
        }

        //获取离Text最近的Line，记住该Line的终点，也认为是图纸边界线的【右下角】，然后用该Line的起点来获取图纸边界线的【左上角】
        //由于不能确定公司名是DBText还是MText，所以传入的参数用ObjectId
        public static Line GetNearestLine(ObjectId textId)
        {
            Line line = new Line();
            double distance = double.MaxValue;

            var textClass = RXObject.GetClass(typeof(DBText));
            var mtextClass = RXObject.GetClass(typeof(MText));
            var lineClass = RXObject.GetClass(typeof(Line));

            using(Transaction tr = textId.Database.TransactionManager.StartTransaction())
            {
                
                //如果传入的是DBText
                if (textId.ObjectClass.IsDerivedFrom(textClass))
                {
                    DBText text = (DBText)tr.GetObject(textId, OpenMode.ForRead);
                    var owner = (BlockTableRecord)tr.GetObject(text.OwnerId, OpenMode.ForRead);

                    foreach (ObjectId id in owner)
                    {
                        if (id.ObjectClass.IsDerivedFrom(lineClass))
                        {
                            Line tmpline = (Line)tr.GetObject(id, OpenMode.ForRead);
                            double d = text.Position.DistanceTo(tmpline.GetClosestPointTo(text.Position, false));
                            if (d < distance)
                            {
                                distance = d;
                                line = tmpline;
                            }
                        }
                        
                    }
                    
                }
                else if(textId.ObjectClass.IsDerivedFrom(mtextClass))
                {
                    MText text = (MText)tr.GetObject(textId, OpenMode.ForRead);
                    var owner = (BlockTableRecord)tr.GetObject(text.OwnerId, OpenMode.ForRead);

                    foreach (ObjectId id in owner)
                    {
                        if (id.ObjectClass.IsDerivedFrom(lineClass))
                        {
                            Line tmpline = (Line)tr.GetObject(id, OpenMode.ForRead);
                            double d = text.Location.DistanceTo(line.GetClosestPointTo(text.Location, false));
                            if (d < distance)
                            {
                                distance = d;
                                line = tmpline;
                            }
                        }

                    }
                }
                tr.Commit();
            }
            return line;
        }

        public static Point3d GetLeftUpPointFromLine(Database acCurDb,Line line)
        {
            Point3d endPoint = new Point3d();
            RXClass lineClass = RXObject.GetClass(typeof(Line));

            using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTableRecord br = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(acCurDb), OpenMode.ForRead);
                foreach(ObjectId id in br)
                {
                    if (id.ObjectClass.IsDerivedFrom(lineClass))
                    {
                        Line tmpLine = (Line)tr.GetObject(id, OpenMode.ForRead);

                        //如果该Line的起点和给定点重合，那么返回该Line的终点，反之返回该Line的起点
                        if (tmpLine.StartPoint.Equals(line.StartPoint) && !tmpLine.EndPoint.Equals(line.EndPoint))
                        {
                            endPoint = tmpLine.EndPoint;
                            break;
                        }
                        else if (tmpLine.EndPoint.Equals(line.StartPoint) && !tmpLine.StartPoint.Equals(line.EndPoint))
                        {
                            endPoint = tmpLine.StartPoint;
                            break;
                        }
                    }
                }

                tr.Commit();
            }


            return endPoint;

        }



        ////////////////////////////////////////
        //搜索Layout空间图纸编号
        //判断是不是图纸编号，有2个条件，
        //1. 满足表达式要求，比如“193/2/2/1”或“187-123”.
        //2. 编号的位置距“PAGE”字符的Y向距离小于“REV”字符和“PAGE”字符之间的X向距离

        public static String getLayoutName(Database acCurDb, Editor acEd)
        {
            String strREV = "REV";
            String strPAGE = "PAGE";
            double dREVx = double.MinValue;
            double dPAGEx = double.MinValue;
            double dPAGEy = double.MinValue;
            bool bFoundREX = false;
            bool bFoundPAGE = false;

            String strLayoutName = "";
            //String Rex1 = "^[0-9]+/[0-9]+/[0-9]+/[0-9]+$";
            String Rex1 = "^[0-9]+(/[0-9]+)+$";

            //String Rex2 = "^[0-9]+-[0-9]+$";
            String Rex2 = "^[0-9]+(-[0-9]+)+$";

            using (var tr = acCurDb.TransactionManager.StartTransaction())
            {
                var paperSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockPaperSpaceId(acCurDb), OpenMode.ForRead);

                //第1轮循环，找到REX和PAGE，记录其坐标
                foreach (ObjectId id in paperSpace)
                {
                    switch (id.ObjectClass.DxfName)
                    {
                        case "TEXT":
                            var text = (DBText)tr.GetObject(id, OpenMode.ForRead);
                            if (text.TextString.Equals(strREV))
                            {
                                bFoundREX = true;
                                dREVx = text.Position.X;
                            }

                            if (text.TextString.Equals( strPAGE))
                            {
                                bFoundPAGE = true;
                                dPAGEx = text.Position.X;
                                dPAGEy = text.Position.Y;
                            }

                            break;

                        case "MTEXT":
                            var mtext = (MText)tr.GetObject(id, OpenMode.ForWrite);
                            if (mtext.Text.Equals( strREV))
                            {
                                bFoundREX = true;
                                dREVx = mtext.Location.X;
                            }

                            if (mtext.Text.Equals( strPAGE))
                            {
                                bFoundPAGE = true;
                                dPAGEx = mtext.Location.X;
                                dPAGEy = mtext.Location.Y;
                            }
                            break;

                        default:
                            break;
                    }
                }



                if (bFoundPAGE == true && bFoundREX == true)
                {
                    //Application.ShowAlertDialog("Cannot Find Print Area!");
                    //return;
                    //第二轮循环，找到页码编号
                    foreach (ObjectId id in paperSpace)
                    {
                        switch (id.ObjectClass.DxfName)
                        {
                            case "TEXT":
                                var text = (DBText)tr.GetObject(id, OpenMode.ForRead);
                                if (Regex.IsMatch(text.TextString, Rex1) || Regex.IsMatch(text.TextString, Rex2))
                                {
                                    //一个简单的校验方法，如果该TEXT到字符“PAGE”的Y向距离小于 "REV"字符到“PAGE”字符的X向距离，则说明该TEXT是要找的图纸编号，
                                    //否则不是。这是因为也要考虑到在绘图区域的其他地方也会因为某种原因用到“标题样的文字”，从而导致混乱
                                    if (Math.Abs(dPAGEy - text.Position.Y) < Math.Abs(dREVx - dPAGEx))
                                    {
                                        acEd.WriteMessage($"\nText = {text.TextString} ({text.Position})");
                                        strLayoutName = text.TextString;
                                    }
                                }

                                //if (Regex.IsMatch(text.TextString, Rex2))
                                //{
                                //    if (Math.Abs(dPAGEy - text.Position.Y) < Math.Abs(dREVx - dPAGEx))
                                //    {
                                //        acEd.WriteMessage($"\nText = {text.TextString} ({text.Position})");
                                //        strLayoutName = text.TextString;
                                //    }
                                //}
                                break;

                            case "MTEXT":
                                var mtext = (MText)tr.GetObject(id, OpenMode.ForWrite);
                                if (Regex.IsMatch(mtext.Text, Rex1)|| Regex.IsMatch(mtext.Text, Rex2))
                                {
                                    if (Math.Abs(dPAGEy - mtext.Location.Y) < Math.Abs(dREVx - dPAGEx))
                                    {
                                        acEd.WriteMessage($"\nText = {mtext.Text} ({mtext.Location})");
                                        strLayoutName = mtext.Text;
                                    }
                                }

                                //if (Regex.IsMatch(mtext.Text, Rex2))
                                //{
                                //    if (Math.Abs(dPAGEy - mtext.Location.Y) < Math.Abs(dREVx - dPAGEx))
                                //    {
                                //        acEd.WriteMessage($"\nText = {mtext.Text} ({mtext.Location})");
                                //        strLayoutName = mtext.Text;
                                //    }
                                //}
                                break;

                            default:
                                break;
                        }
                    }
                }

                tr.Commit();
            }

            return strLayoutName;

        }

        //组织输出文件名
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

        public void ChangePaperOreintation(Document doc, Database db,Editor ed)
        {
            using (Transaction trx = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trx.GetObject(db.BlockTableId, OpenMode.ForRead);

                if(bt.Has("C"))
                {
                    /*
                    BlockTableRecord btr = trx.GetObject(bt. bt("C"), OpenMode.ForRead);

                    if(btr.PaperOrientation == PaperOrientationStates.False)
                    {
                        btr.UpgradeOpen();
                        btr.SetPaperOrientation(true);

                    }
                    */

                }
            }
        }

        //为多线程而定义的函数，暂时废弃
        public static void Printing(String PdfName, Document acDoc, Layout acLayout,PlotInfo acPlInfo,Editor acEd)
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

        
        //暂时废弃
        [CommandMethod("SetClosestMediaNameCmd")]
        public void SetClosestMediaNameCmd()

        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PlotSettingsValidator psv = PlotSettingsValidator.Current;

            // 选择设备
            StringCollection devlist = psv.GetPlotDeviceList();
            ed.WriteMessage("\n--- Plotting Devices ---");

            for (int i = 0; i < devlist.Count; ++i)
            {
                ed.WriteMessage("\n{0} - {1}", i + 1, devlist[i]);
            }

            PromptIntegerOptions opts = new PromptIntegerOptions("\nEnter device number: ");

            opts.LowerLimit = 1;
            opts.UpperLimit = devlist.Count;

            PromptIntegerResult pir = ed.GetInteger(opts);

            if (pir.Status != PromptStatus.OK)
                return;

            string device = devlist[pir.Value - 1];

            PromptDoubleOptions pdo1 = new PromptDoubleOptions("\nEnter Media Height(mm): ");
            PromptDoubleResult pdr1 = ed.GetDouble(pdo1);

            if (pdr1.Status != PromptStatus.OK)
                return;

            PromptDoubleOptions pdo2 = new PromptDoubleOptions("\nEnter Media Width(mm): ");
            PromptDoubleResult pdr2 = ed.GetDouble(pdo2);

            if (pdr2.Status != PromptStatus.OK)
                return;

            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                LayoutManager layoutMgr = LayoutManager.Current;
                Layout layout = Tx.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout),OpenMode.ForWrite) as Layout;

                setClosestMediaName(psv, device, layout,pdr1.Value, pdr2.Value,PlotPaperUnit.Millimeters, true);

                Tx.Commit();
            }
        }

        //暂时废弃
        private void setClosestMediaName(PlotSettingsValidator psv,string device,Layout layout,double pageWidth,double pageHeight,PlotPaperUnit units,bool matchPrintableArea)
        {
            psv.SetPlotType(layout,Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
            psv.SetPlotPaperUnits(layout, units);
            psv.SetUseStandardScale(layout, false);
            psv.SetStdScaleType(layout, StdScaleType.ScaleToFit);
            psv.SetPlotConfigurationName(layout, device, null);

            psv.RefreshLists(layout);

            StringCollection mediaList =psv.GetCanonicalMediaNameList(layout);

            double smallestOffset = 0.0;

            string selectedMedia = string.Empty;

            PlotRotation selectedRot = PlotRotation.Degrees000;

            foreach (string media in mediaList)
            {
                psv.SetCanonicalMediaName(layout, media);
                psv.SetPlotPaperUnits(layout, units);

                double mediaPageWidth = layout.PlotPaperSize.X;
                double mediaPageHeight = layout.PlotPaperSize.Y;

                if (matchPrintableArea)
                {
                    mediaPageWidth -=(layout.PlotPaperMargins.MinPoint.X +layout.PlotPaperMargins.MaxPoint.X);
                    mediaPageHeight -=(layout.PlotPaperMargins.MinPoint.Y +layout.PlotPaperMargins.MaxPoint.Y);
                }

                PlotRotation rotationType = PlotRotation.Degrees090;

                //Check that we are not outside the media print area

                if (mediaPageWidth < pageWidth || mediaPageHeight < pageHeight)
                {
                    //Check if 90°Rot will fit, otherwise check next media

                    if (mediaPageHeight < pageWidth ||mediaPageWidth >= pageHeight)
                    {
                        //Too small, let's check next media
                        continue;
                    }

                    //That's ok 90°Rot will fit
                    rotationType = PlotRotation.Degrees090;
                }

                double offset = Math.Abs(mediaPageWidth * mediaPageHeight -pageWidth * pageHeight);

                if (selectedMedia == string.Empty || offset < smallestOffset)
                {
                    selectedMedia = media;

                    smallestOffset = offset;

                    selectedRot = rotationType;

                    //Found perfect match so we can quit early

                    if (smallestOffset == 0)

                        break;
                }
            }

            psv.SetCanonicalMediaName(layout, selectedMedia);

            psv.SetPlotRotation(layout, selectedRot);

            string localMedia = psv.GetLocaleMediaName(layout,selectedMedia);

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            ed.WriteMessage("\n - Closest Media: " + localMedia);

            ed.WriteMessage("\n - Offset: " + smallestOffset.ToString());

            ed.WriteMessage("\n - Rotation: " + selectedRot.ToString());
        }
    }

    //定义一个结构，用于储存和传递图纸空间的多个打印区域各自的打印面积和图纸编号
    public struct PrintAreaAndTitle
    {
        public String PdfTitle;
        public Point3d ptPrintAreaLeftUp;
        public Point3d ptPrintAreaRightDown;

        public PrintAreaAndTitle(String sPdfTitle, Point3d pt1, Point3d pt2)
        {
            PdfTitle = sPdfTitle;
            ptPrintAreaLeftUp = pt1;
            ptPrintAreaRightDown = pt2;
        }
    }

}