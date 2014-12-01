using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace Career_Fair.model
{
    class Session
    {
        private Presenter sessionPresenter;
        
        List<Student> sessionOneStudentList, sessionTwoStudentList, sessionThreeStudentList;

        public int Capacity
        {
            get { return sessionPresenter.PresenterRoom.RoomCapacity; }
        }

        public Presenter SessionPresenter
        {
            get { return sessionPresenter; }
            set { sessionPresenter = value; }
        }

        public List<Student> SessionOneList
        {
            get { return sessionOneStudentList; }
            set { sessionOneStudentList = value; }
        }

        public List<Student> SessionTwoList
        {
            get { return sessionTwoStudentList; }
            set { sessionTwoStudentList = value; }
        }

        public List<Student> SessionThreeList
        {
            get { return sessionThreeStudentList; }
            set { sessionThreeStudentList = value; }
        }

        public int SessionOneCount
        {
            get { return sessionOneStudentList.Count; }
        }

        public int SessionTwoCount
        {
            get { return sessionTwoStudentList.Count; }
        }

        public int SessionThreeCount
        {
            get { return sessionThreeStudentList.Count; }
        }
        
        public Session(Presenter sessionPresenter)
        {
            this.sessionPresenter = sessionPresenter;
            
            sessionOneStudentList = new List<Student>();
            sessionTwoStudentList = new List<Student>();
            sessionThreeStudentList = new List<Student>();
        }
    }
}
