using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using Career_Fair.model;
using Career_Fair.controller;

namespace Career_Fair.view
{
    public partial class CareerFairGUI : Form
    {

        private WorkshopData currentData;
        private WorkshopPrep prepData;
        private string workshopDataPath, studentDataPath;
        private List<District> districtChangeList;
        private WorkshopIO fileIO;

        public String StatusText
        {
            get { return statusLabel.Text;  }
            set { statusLabel.Text = value; }
        }

        public string WorkshopDataFilePath
        {
            get { return workshopDataPath; }
        }

        public string StudentDataFilePath
        {
            get { return studentDataPath; }
        }


        public CareerFairGUI()
        {
            InitializeComponent();
            prepData = new WorkshopPrep();
            currentData = new WorkshopData();
            districtChangeList = new List<District>();
            fileIO = new WorkshopIO();
            
        }

       
        private void workshopFileButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;

            DialogResult userChoice = openFileDialog1.ShowDialog();
            if (userChoice == DialogResult.OK)
            {
                workshopDataPath = openFileDialog1.FileName;
                currentData.WorkshopInfoFilePath = WorkshopDataFilePath;
                chooseStudentRequestButton.Enabled = true;
            }
            
        }

        private void chooseStudentRequestButton_Click(object sender, EventArgs e)
        {   
            openFileDialog1.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;

            DialogResult userChoice = openFileDialog1.ShowDialog();
            if (userChoice == DialogResult.OK)
            {
                generateScheduleButton.Enabled = true;
                studentDataPath = openFileDialog1.FileName;
                currentData.RequestFilePath = StudentDataFilePath;
            }
        }

        private void generateScheduleButton_Click(object sender, EventArgs e)
        {
            currentData.startSchedule();
            statusLabel.Text = currentData.StatusText;
            currentData.verifyPopularity();
            statusLabel.Text = currentData.StatusText;
            
            currentData.generateScheduleFromLists();
            statusLabel.Text += currentData.StatusText;
            currentData.createWorkshopExportExcel();
            if (currentData.ScheduleOK)
            {
                launchScheduleButton.Enabled = true;
                launchScheduleButton.Visible = true;
            }
        }

        private void openSchedule()
        {
            string filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Schedule.xlsx";
            System.Diagnostics.Process.Start(filePath);
        }

        private void launchScheduleButton_Click(object sender, EventArgs e)
        {
            openSchedule();
        }

        private void aboutSchedulerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About wfcAbout = new About();
            wfcAbout.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void helpFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFilesX86) + "\\Career Fair\\Career Fair help.pdf";
            System.Diagnostics.Process.Start(filePath);
        }

        private void loadRequestsButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);

            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;

            DialogResult userChoice = openFileDialog1.ShowDialog();
            if (userChoice == DialogResult.OK)
            {
                studentDataPath = openFileDialog1.FileName;
                prepData.WorkshopDataPath = studentDataPath;
                fileIO.WorkshopDataPath = studentDataPath;
            }
            prepData.processRequests(fileIO.readStudentData());
            prepData.prepareExport();
            prepData.exportSuggestedSchedule();
            prepData.exportSampleDataFile();
            openFiles();
        }

        private void openFiles()
        {
            string filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Data File.xlsx";
            System.Diagnostics.Process.Start(filePath);

            filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Career Fair Suggestions.xlsx";
            System.Diagnostics.Process.Start(filePath);
        }
        
    }
}
