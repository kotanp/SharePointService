﻿using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using _Application = Microsoft.Office.Interop.Word._Application;
using Word = Microsoft.Office.Interop.Word;

namespace SharePointService.Service
{
    public class ConverterService : IConverterService
    {

        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        public static Random random = new Random();

        // PDF conversion attributes
        private object oMissing = System.Reflection.Missing.Value;
        private object isVisible = true;
        private object readOnly = true;

        private const string PDF_EXTENSTION = ".pdf";
                                    

        public string ConvertToPdf(byte[] docItem, string fileExtension)
        {
            /*            var tmpFile = Path.GetTempFileName();
                        var tmpFileStream = File.OpenWrite(tmpFile);
                        tmpFileStream.Write(docItem, 0, docItem.Length);*/
            string newFileExtension = fileExtension;
            if (!fileExtension.Contains("."))
            {
                newFileExtension = "." + fileExtension;
            }

            byte[] pdfBytes = null;
            if (fileExtension.Equals("docx") || fileExtension.Equals("doc") || fileExtension.Equals(".docx") || fileExtension.Equals(".doc"))
            {
                pdfBytes = ConvertDocToPDf(docItem, newFileExtension);
            }
            else if (fileExtension.Equals("xlsx") || fileExtension.Equals("xls") || fileExtension.Equals(".xlsx")
                || fileExtension.Equals(".xls"))
            {
                pdfBytes = ConvertXlsToPdf(docItem, newFileExtension);
            }

            return JsonConvert.SerializeObject(pdfBytes);
        }

        /*
         * Converts the docx byte to pdf bytes
         */
        private byte[] ConvertDocToPDf(byte[] docItem, string newFileExtension)
        {
            string tempFileName = CreateTempFile(newFileExtension, docItem);

            _Application _app = new Word.Application
            {

                // Make this instance of word invisible (Can still see it in the taskmgr).
                Visible = false
            };
            _Document doc = _app.Documents.Open(tempFileName, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();
            string pdfFileName = GetPdfFileName();

            doc.SaveAs(pdfFileName, WdSaveFormat.wdFormatPDF);
            doc.Close(false);
            _app.Quit(ref oMissing, ref oMissing, ref oMissing);

            byte[] pdfBytes = File.ReadAllBytes(pdfFileName);

            File.Delete(tempFileName);
            File.Delete(pdfFileName);

            _app = null;

            return pdfBytes;
        }

        /*
         * Converts the xlsx to pdf bytes
         */
        private byte[] ConvertXlsToPdf(byte[] docItem, string newFileExtension)
        {
            string tempFileName = CreateTempFile(newFileExtension, docItem);

            Microsoft.Office.Interop.Excel.Application _app = new Microsoft.Office.Interop.Excel.Application {
                Visible = false
            };

            Workbook workbook = _app.Workbooks.Open(tempFileName);
            workbook.Activate();
            string pdfFileName = GetPdfFileName();
            
            workbook.ExportAsFixedFormat2(XlFixedFormatType.xlTypePDF, pdfFileName);
            workbook.Close(0);
            _app.Quit();

            byte[] pdfBytes = File.ReadAllBytes(pdfFileName);

            File.Delete(tempFileName);
            File.Delete(pdfFileName);

            _app = null;

            return pdfBytes;

        }

        /*
         * Generate random pdf fileName
         */
        private string GenerateRandomFileName()
        {
            return new string(Enumerable.Repeat(chars, 5).Select(s => s[random.Next(s.Length)]).ToArray());
        }

        /*
         * Returns the new temp file name, used for the conversion
         */
        private string GetTempFileName(string newFileExtension)
        {
            return Directory.GetCurrentDirectory() + "\\" + GenerateRandomFileName() + newFileExtension;
        }

        /*
         * Returns the new pdf file path
         */
        private string GetPdfFileName()
        {
            return Directory.GetCurrentDirectory() + "\\" + GenerateRandomFileName() + PDF_EXTENSTION;
        }

        /*
         * Create a temp file and returns it's name
         */
        private string CreateTempFile(string newFileExtension, byte[] docItem)
        {
            string tempFileName = GetTempFileName(newFileExtension);
            File.WriteAllBytes(tempFileName, docItem);
            return tempFileName;
        }
    }
}