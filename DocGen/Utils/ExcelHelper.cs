﻿using DocGen.View.Blank;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocGen.Utils
{
    static class ExcelHelper
    {
        private static Excel.Application xlApp = (Excel.Application)Globals.ThisAddIn.Application;

        public static void DisableUpdating()
        {
            xlApp.ScreenUpdating = false;
            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            xlApp.EnableEvents = false;
            xlApp.DisplayStatusBar = false;
        }

        public static void EnableUpdating()
        {
            xlApp.ScreenUpdating = true;
            xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            xlApp.EnableEvents = true;
            xlApp.DisplayStatusBar = true;
        }

        public static void SetPrintSettings(Excel.Worksheet sheet)
        {
            xlApp.PrintCommunication = false;
            Excel.PageSetup pageSetup = (Excel.PageSetup)sheet.PageSetup;

            pageSetup.BottomMargin = xlApp.CentimetersToPoints(0.0);
            pageSetup.TopMargin = xlApp.CentimetersToPoints(0.5);
            pageSetup.LeftMargin = xlApp.CentimetersToPoints(0.7);
            pageSetup.RightMargin = xlApp.CentimetersToPoints(0.5);
            pageSetup.HeaderMargin = xlApp.CentimetersToPoints(0.0);
            pageSetup.FooterMargin = xlApp.CentimetersToPoints(0.0);
            pageSetup.PrintArea = "A:Z";
            pageSetup.Zoom = 100;
            xlApp.PrintCommunication = true;
            SetPageSpecificSettings(sheet, pageSetup);
            
        }

        public static void SetPageSpecificSettings(Excel.Worksheet sheet, 
                                                    Excel.PageSetup pageSetup)
        {
            string type = ListPage.GetDocumentType(sheet);
            switch (type)
            {
                case "Перечень элементов":
                    pageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                    pageSetup.PrintArea = "A:Z";
                    break;
                case "Спецификация":
                    pageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                    pageSetup.PrintArea = "A:Z";
                    break;
                case "Ведомость покупных изделий":
                    pageSetup.PaperSize = Excel.XlPaperSize.xlPaperA3;
                    pageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    pageSetup.PrintArea = "A:AT";
                    break;
                default:
                    break;
            }
        }

        public static void DisableZeros()
        {
            xlApp.ActiveWindow.DisplayZeros = false;
        }

        public static void SetupPageBreaksView()
        {
            Excel.Window window = ((Excel.Window)Globals.ThisAddIn.Application.ActiveWindow);
            window.View = Excel.XlWindowView.xlPageBreakPreview;

        }

        public static void SetupNormalView()
        {
            Excel.Window window = ((Excel.Window)Globals.ThisAddIn.Application.ActiveWindow);
            window.View = Excel.XlWindowView.xlNormalView;
        }
    }
}
