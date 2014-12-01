
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Windows.Forms;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

namespace Career_Fair.model
{
    /// <summary>
    /// Creates an excel file to determine the number of sessions needed based on student requests.
    /// This will be executed before the scheduling component of the application.
    /// </summary>
    class WorkshopPrep
    {
        private List<Student> studentList;
        private List<Request> requestList;
        private List<Presenter> presenterList;
        private List<Room> roomList;
        private List<Session> sessionList;
        private List<Prep> prepList;
        
        private class Prep
        {
            public String presenterName;
            public int firstCount, secondCount, thirdCount, suggestedRooms;
        }

        private String workshopDataPath;


        public String WorkshopDataPath
        {
            get { return workshopDataPath; }
            set { workshopDataPath = value; }
        }


        #region Prepare Suggestions
        private void buildPrepList()
        {
            prepList = new List<Prep>();
            
            foreach (Request currentRequest in requestList)
            {
                Prep tempPrepOne, tempPrepTwo, tempPrepThree, tempPrepFour, tempPrepFive;


                tempPrepOne = prepList.Find(delegate(Prep curr) { return currentRequest.RequestOne.Equals(curr.presenterName); });
                tempPrepTwo = prepList.Find(delegate(Prep curr) { return currentRequest.RequestTwo.Equals(curr.presenterName); });
                tempPrepThree = prepList.Find(delegate(Prep curr) { return currentRequest.RequestThree.Equals(curr.presenterName); });
                tempPrepFour = prepList.Find(delegate(Prep curr) { return currentRequest.RequestFour.Equals(curr.presenterName); });
                tempPrepFive = prepList.Find(delegate(Prep curr) { return currentRequest.RequestFive.Equals(curr.presenterName); });
               
                 
                 
                if (tempPrepOne == null)
                {
                    tempPrepOne = new Prep();
                    tempPrepOne.presenterName = currentRequest.RequestOne;
                    prepList.Add(tempPrepOne);
                }
                if (tempPrepTwo == null)
                {
                    tempPrepTwo = new Prep();
                    tempPrepTwo.presenterName = currentRequest.RequestTwo;
                    prepList.Add(tempPrepTwo);
                }
                if (tempPrepThree == null)
                {
                    tempPrepThree = new Prep();
                    tempPrepThree.presenterName = currentRequest.RequestThree;
                    prepList.Add(tempPrepThree);
                }
                if (tempPrepFour == null)
                {
                    tempPrepFour = new Prep();
                    tempPrepFour.presenterName = currentRequest.RequestFour;
                    
                    prepList.Add(tempPrepFour);
                }
                if (tempPrepFive == null)
                {
                    tempPrepFive = new Prep();
                    tempPrepFive.presenterName = currentRequest.RequestFive;
                    
                    prepList.Add(tempPrepFive);
                }
            }
        }

        private void buildCountRequests()
        {
            foreach (Request currentRequest in requestList)
            {
                Prep tempPrepOne = prepList.Find(delegate(Prep curr) { return curr.presenterName.Equals(currentRequest.RequestOne); });
                Prep tempPrepTwo = prepList.Find(delegate(Prep curr) { return curr.presenterName.Equals(currentRequest.RequestTwo); });
                Prep tempPrepThree = prepList.Find(delegate(Prep curr) { return curr.presenterName.Equals(currentRequest.RequestThree); });

                tempPrepOne.firstCount++;
                tempPrepTwo.secondCount++;
                tempPrepThree.thirdCount++;
            }
        }

        private void calculateRecommendations()
        {
            foreach (Prep currentPrep in prepList)
            {
                int totalRequests = (currentPrep.firstCount + currentPrep.secondCount + currentPrep.thirdCount);
                if (totalRequests < 10)
                {
                    currentPrep.suggestedRooms = 0;
                }
                else if ((totalRequests / 35) <3)
                {
                    currentPrep.suggestedRooms = 1;
                }
                else if ((totalRequests / 35) < 5)
                {
                    currentPrep.suggestedRooms = 2;
                }
                else
                {
                    currentPrep.suggestedRooms = 3;
                }
            }
        }

        public void prepareExport()
        {
            buildPrepList();
            buildCountRequests();
            calculateRecommendations();
        }
        #endregion

        #region Reading Excel Data

        public void processRequests(ExcelPackage currentPackage)
        {
            ExcelWorkbook currentWorkbook = currentPackage.Workbook;
            if (currentWorkbook != null)
            {
                ExcelWorksheet requestSheet = currentWorkbook.Worksheets[1];
                createStudentAndRequestList(requestSheet);
            }
        }

        #endregion

        #region Creating Excel data for recommendation.

        public void exportSampleDataFile()
        {
            FileInfo currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Data File.xlsx");
            if (currentFile.Exists)
            {
                currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Data File (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                if (currentFile.Exists)
                {
                    currentFile.Delete();
                    currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Data File (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                }
            }

            using (ExcelPackage currentExcel = new ExcelPackage(currentFile))
            {
                ExcelWorksheet currentSheet = currentExcel.Workbook.Worksheets.Add("Rooms");

                currentSheet = currentExcel.Workbook.Worksheets.Add("Presenters");
                int currentRowCounter = 2;

                currentSheet.Cells["A1"].Value = "Presenter";
                currentSheet.Cells["B1"].Value = "Description";
                currentSheet.Cells["C1"].Value = "Room";

                for (int currentCount = 0; currentCount < prepList.Count; currentCount++)
                {
                    Prep currentPrep = prepList[currentCount];
                    for (int prepRow = 0; prepRow < currentPrep.suggestedRooms; prepRow++)
                    {
                        currentSheet.Cells["A" + currentRowCounter].Value = currentPrep.presenterName + "-" + (prepRow + 1);
                        currentRowCounter++;
                    }
                }

                currentExcel.Save();
            }
        }

        public void exportSuggestedSchedule()
        {
            FileInfo currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Suggestions.xlsx");
            if (currentFile.Exists)
            {
                currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Suggestions (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                if (currentFile.Exists)
                {
                    currentFile.Delete();
                    currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Suggestions (" + DateTime.Today.Month + "-" + DateTime.Today.Day + ").xlsx");
                }
            }

            int columnCount = 5;
            using (ExcelPackage currentExcel = new ExcelPackage(currentFile))
            {
                ExcelWorksheet currentSheet = currentExcel.Workbook.Worksheets.Add("Session Needs");
                int currentRowCounter = 2;

                currentSheet.Cells["A1"].Value = "Session Title";
                currentSheet.Cells["B1"].Value = "Session First Level Requests";
                currentSheet.Cells["C1"].Value = "Session Second Level Requests";
                currentSheet.Cells["D1"].Value = "Session Third Level Requests";
                currentSheet.Cells["E1"].Value = "Estimated Rooms needed for session";

                currentSheet.Cells["A1"].AutoFitColumns();

                String headerRange = "A1:" + Convert.ToChar('A' + columnCount - 1) + 1;
                formatExcelHeader(headerRange, currentSheet);

                for (int count = 0; count < prepList.Count; count++)
                {
                    currentSheet.Cells[currentRowCounter, 1].Value = prepList[count].presenterName;
                    currentSheet.Cells[currentRowCounter, 2].Value = prepList[count].firstCount;
                    currentSheet.Cells[currentRowCounter, 3].Value = prepList[count].secondCount;
                    currentSheet.Cells[currentRowCounter, 4].Value = prepList[count].thirdCount;
                    currentSheet.Cells[currentRowCounter, 5].Value = prepList[count].suggestedRooms;
                    currentRowCounter++;
                }

                currentExcel.Save();
            }
        }

        #endregion

        
        private void fillPresenterSheet(ExcelWorksheet presenterSheet)
        {
            presenterSheet.Cells["A1"].Value = "Presenter Title";
            presenterSheet.Cells["B1"].Value = "Presenter Description";
            presenterSheet.Cells["C1"].Value = "Presenter Room";
            presenterSheet.Cells["A1"].AutoFitColumns();

            String headerRange = "A1:C1";
            formatExcelHeader(headerRange, presenterSheet);
        }

        private void formatExcelHeader(String headerRange, ExcelWorksheet currentSheet)
        {
            using (ExcelRange currentRange = currentSheet.Cells[headerRange])
            {
                currentRange.Style.WrapText = false;
                currentRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                currentRange.Style.Font.Bold = true;
                currentRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                currentRange.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                currentRange.Style.Font.Color.SetColor(Color.White);

            }
        }

        private void createStudentAndRequestList(ExcelWorksheet requestSheet)
        {

            string firstName, lastName, requestOne, requestTwo, requestThree, requestFour, requestFive, studentTeacher, errorStudent;
            int studentID=0;
            //DateTime studentTime;
            errorStudent = "";
            studentList = new List<Student>();
            requestList = new List<Request>();

            try
            {
                for (int row = 2; row <= requestSheet.Dimension.End.Row; row++)
                {
                    studentID = row;
                    
                    firstName = (String)requestSheet.Cells[row, 1].Value.ToString();
                    lastName = (String)requestSheet.Cells[row, 2].Value.ToString();
                    errorStudent = lastName + ", " + firstName;

                    //Check for varying date types
                    //double serialDate = double.Parse(requestSheet.Cells[row, 1].Value.ToString());
                    //studentTime = DateTime.FromOADate(serialDate);

                    studentTeacher = (String)requestSheet.Cells[row, 3].Value.ToString();
                    requestOne = (String)requestSheet.Cells[row, 4].Value.ToString();
                    requestTwo = (String)requestSheet.Cells[row, 5].Value.ToString();
                    requestThree = (String)requestSheet.Cells[row, 6].Value.ToString();
                    requestFour = (String)requestSheet.Cells[row, 7].Value.ToString();
                    requestFive = (String)requestSheet.Cells[row, 8].Value.ToString();

                    Request currentRequest;
                    Student currentStudent = new Student(firstName, lastName, studentTeacher);
                    currentRequest = new Request(currentStudent, requestOne, requestTwo, requestThree, requestFour, requestFive);
                    currentStudent.StudentRequest = currentRequest;

                    requestList.Add(currentRequest);
                    studentList.Add(currentStudent);

                }

            }
            catch (Exception currentException)
            {
                MessageBox.Show("Last successful imported student was: " + errorStudent + " at row: " + studentID + " check the format of the fields in the student request file.");
                MessageBox.Show("\n\nReport this if fixing the request file does not solve the problem.\n\n"+currentException.Message);
            }
        }


        #region Create Lists

        /// <summary>
        /// Creates the room list
        /// </summary>
        /// <param name="roomSheet">Extracted Excel 2007+ data</param>
        private void createRoomList(ExcelWorksheet roomSheet)
        {
            roomList = new List<Room>();
            String roomName, errorRoom;
            int roomCapacity;
            errorRoom = "";

            try
            {
                for (int row = 2; row <= roomSheet.Dimension.End.Row; row++)
                {
                    roomName = (String)roomSheet.Cells[row, 1].Value.ToString();
                    errorRoom = roomName;
                    roomCapacity = Convert.ToInt32(roomSheet.Cells[row, 2].Value);
                    roomList.Add(new Room(roomName, roomCapacity));
                    
                }
            }
            catch (Exception currentException)
            {
                MessageBox.Show("last successful room was: " + errorRoom + " check the .XLSX file");
                MessageBox.Show("\n\nReport this if fixing the Data file does not solve the problem.\n\n" + currentException.Message);
            }




        }

        /// <summary>
        /// Creates the presenter list.  Must be completed after the room list as it uses that for a search for specific room object
        /// </summary>
        /// <param name="presentersSheet"></param>
        private void createPresenterList(ExcelWorksheet presentersSheet)
        {
            presenterList = new List<Presenter>();

            String PresenterTitle, presenterRoom, errorPresenter;
            errorPresenter = "";
            try
            {
                for (int row = 2; row <= presentersSheet.Dimension.End.Row; row++)
                {
                    PresenterTitle = (String)presentersSheet.Cells[row, 1].Value;
                    errorPresenter = PresenterTitle;
                    presenterRoom = (String)presentersSheet.Cells[row, 2].Value.ToString();

                    Room currentRoom = roomList.Find(delegate(Room current) { return current.RoomName.Equals(presenterRoom); });
                    presenterList.Add(new Presenter(PresenterTitle, currentRoom));
                    
                }
            }
            catch (Exception currentException)
            {
                MessageBox.Show("last successful presenter was: " + errorPresenter + " check the .XLSX file");
                MessageBox.Show("\n\nReport this if fixing the Data file does not solve the problem.\n\n" + currentException.Message);
            }
        }

        

        /// <summary>
        /// Creates the session list.  Must occur after the presenter list is created.
        /// </summary>
        private void createSessionList()
        {
            sessionList = new List<Session>();
            foreach (Presenter currentPresenter in presenterList)
            {
                sessionList.Add(new Session(currentPresenter));
            }

        }

        #endregion



    }
}
