using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace Career_Fair.model
{
    public class District
    {
        private ArrayList schoolList;
        private string districtName;
        private int districtPercentage;
        private int totalStudents;
        private int maxStudents;


        public int MaxStudents
        {
            get { return maxStudents; }
            set { maxStudents = value; }
        }

        public int TotalStudents
        {
            get { return totalStudents; }
            set { totalStudents = value; }
        }

        public string DistrictName
        {
            get { return districtName; }
            set { districtName = value; }
        }

        public int DistrictPercentage
        {
            get { return districtPercentage; }
            set { districtPercentage = value; }
        }

        public District(string districtName, int districtPercentage, int maxStudents)
        {
            this.districtName = districtName;
            this.districtPercentage = districtPercentage;
            this.maxStudents = maxStudents;
            totalStudents = 0;
        }
    }
}
