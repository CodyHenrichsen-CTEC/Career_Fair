using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Career_Fair.model
{
    class School
    {
        private District schoolDistrict;
        private String schoolName;

        public District SchoolDistrict
        {
            get { return schoolDistrict; }
            set { schoolDistrict = value; }
        }

        public String SchoolName
        {
            get { return schoolName; }
            set { schoolName = value; }
        }

        public School(District schoolDistrict, String schoolName)
        {
            this.schoolDistrict = schoolDistrict;
            this.schoolName = schoolName;
        }
    }
}
