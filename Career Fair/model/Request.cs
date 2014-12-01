using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Career_Fair.model
{
    class Request
    {
        private Student requestingStudent;
        private string requestOne, requestTwo, requestThree, requestFour, requestFive;
        private DateTime requestTime;

        public Student RequestingStudent
        {
            get { return requestingStudent; }
            set { requestingStudent = value; }
        }

        public string RequestOne
        {
            get { return requestOne; }
            set { requestOne = value; }
        }

        public string RequestTwo
        {
            get { return requestTwo; }
            set { requestTwo = value; }
        }

        public string RequestThree
        {
            get { return requestThree; }
            set { requestThree = value; }
        }

        public string RequestFour
        {
            get { return requestFour; }
            set { requestFour = value; }
        }

        public string RequestFive
        {
            get { return requestFive; }
            set { requestFive = value; }
        }

        public DateTime RequestTime
        {
            get { return requestTime; }
            set { requestTime = value; }
        }

        public Request(Student requestingStudent,  string requestOne, string requestTwo, string requestThree, string requestFour, string requestFive)
        {
            this.requestingStudent = requestingStudent;
            this.requestOne = requestOne;
            this.requestTwo = requestTwo;
            this.requestThree = requestThree;
            this.requestFour = requestFour;
            this.requestFive = requestFive;
        }

    }
}
