using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;

namespace HotManchego
{
    class VBAGenerator
    {
        public static void Main(string[] args)
        {
            if(args.Length == 0) {
                 System.Console.WriteLine("Usage: hot-manchego.exe blank.xlsm vba.txt\nThe first argument is a blank XLSM file.\nThe second argument is the VBA you want embedded in the XLSM file.");
                 System.Environment.Exit(1);
            }
            if(ars.Length == 1){
            	System.Console.WriteLine("Usage: hot-manchego.exe blank.xlsm vba.txt\nThe first argument is a blank XLSM file.\nThe second argument is the VBA you want embedded in the XLSM file.");
                System.Environment.Exit(1);
	    	}

	    	var outFile = new FileInfo(Directory.GetCurrentDirectory() + "\\" + args[0]);
            var vbaFile = new FileInfo(Directory.GetCurrentDirectory() + "\\" + args[1]);
            FillVBA(outFile, vbaFile);
          
        }
        private static void FillVBA(FileInfo outFile, FileInfo vbaFile)
        {
            ExcelPackage pck = new ExcelPackage();

            //Add a worksheet.
            var ws = pck.Workbook.Worksheets.Add("Sheet1");
            //ws.Drawings.AddShape("VBASampleRect", eShapeStyle.RoundRect);
            
            //Create a vba project and set password permissions             
            pck.Workbook.CreateVBAProject();
	    	pck.Workbook.VbaProject.Protection.SetPassword("EPPlus");

            //Read in vba code from file
            pck.Workbook.CodeModule.Code = File.ReadAllText(vbaFile.FullName);

            //Optionally, Sign the code with your company certificate.
            //X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //pck.Workbook.VbaProject.Signature.Certificate = store.Certificates[0];

            //Save as xlsm
            pck.SaveAs(outFile);
        }
    }
}

