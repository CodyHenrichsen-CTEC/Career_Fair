using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Career_Fair.model
{
    class DistrictClass
    {
        private String districtName, className;
        private int count, max;

        public int Max
        {
            get { return max; }
            set { max = value; }
        }
        public String DistrictName
        {
            get { return districtName; }
            set { districtName = value; }
        }
        public String ClassName
        {
            get { return className; }
            set { className = value; }
        }
        public int Count
        {
            get { return count; }
            set { count = value; }
        }
        public DistrictClass(String districtName, String className)
        {
            this.districtName = districtName;
            this.className = className;
            count = 0;
        }
    }
}
