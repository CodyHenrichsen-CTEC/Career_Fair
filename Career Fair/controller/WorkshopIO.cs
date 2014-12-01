using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using Career_Fair.model;
using System.Drawing;
using System.Windows.Forms;

namespace Career_Fair.controller
{
    class WorkshopIO
    {
        private Dictionary<String, Int16> firstRequests;
        private Dictionary<String, Int16> secondRequests;

        private String workshopDataPath;


        public String WorkshopDataPath
        {
            get { return workshopDataPath; }
            set { workshopDataPath = value; }
        }

        private bool checkSheetsMatch(ExcelWorkbook current)
        {
            bool sheetsMatch = false;
            int rooms = 0, presenters = 0;

            for (int count = 0; count < 2; count++)
            {
                ExcelWorksheet currentWorksheet = current.Worksheets[count + 1];
                if (currentWorksheet.Name.Equals("Rooms"))
                {
                    rooms = 1;
                }
                if (currentWorksheet.Name.Equals("Presenters"))
                {
                    presenters = 1;
                }
            }
            sheetsMatch = (rooms + presenters == 2);
            return sheetsMatch;

        }


        public ExcelPackage readStudentData()
        {
            FileInfo requestFile = new FileInfo(workshopDataPath);
            ExcelPackage currentExcelFile = new ExcelPackage(requestFile);
            return currentExcelFile;

        }

        public ExcelPackage readWorkshopData()
        {
            return null;
        }
        
        public void writeWorkshopSchedule(ExcelPackage scheduleProject)
        {
            FileInfo currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Schedule.xlsx");
            if (currentFile.Exists)
            {
                currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Schedule (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                if (currentFile.Exists)
                {
                    currentFile.Delete();
                    currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Schedule (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                }
            }

            using (FileStream ioStream = currentFile.Create())
            {
                byte[] excelBytes = scheduleProject.GetAsByteArray();
                ioStream.Write(excelBytes, 0, excelBytes.Length);
            }
        }

        public void writeSampleDataFile(ExcelPackage dataProject)
        {
            FileInfo currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Data.xlsx");
            if (currentFile.Exists)
            {
                currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Data (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                if (currentFile.Exists)
                {
                    currentFile.Delete();
                    currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Data (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                }
            }

            using (FileStream ioStream = currentFile.Create())
            {
                byte[] excelBytes = dataProject.GetAsByteArray();
                ioStream.Write(excelBytes, 0, excelBytes.Length);
            }
        }

        public void writeSuggestedSpeakers(ExcelPackage currentProject)
        {
            FileInfo currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Presenter Needs.xlsx");
            if (currentFile.Exists)
            {
                currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Presenter Needs (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                if (currentFile.Exists)
                {
                    currentFile.Delete();
                    currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Presenter Needs (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                }
            }

            using (FileStream ioStream = currentFile.Create())
            {
                byte[] excelBytes = currentProject.GetAsByteArray();
                ioStream.Write(excelBytes, 0, excelBytes.Length);
            }
        }

        


        }
}
