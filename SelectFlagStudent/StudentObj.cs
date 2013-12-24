using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SelectFlagStudent
{
    class StudentObj
    {
        private string id, className, seatNo, studentNo, name;
        private DateTime time;

        public StudentObj(string id, string className, string seatNo, string studentNo, string name, DateTime time)
        {
            this.id = id;
            this.className = className;
            this.seatNo = seatNo;
            this.studentNo = studentNo;
            this.name = name;
            this.time = time;
        }

        public DateTime Time
        {
            get { return time; }
            set { time = value; }
        }

        public string StudentNo
        {
            get { return studentNo; }
            set { studentNo = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string SeatNo
        {
            get { return seatNo; }
            set { seatNo = value; }
        }

        public string ClassName
        {
            get { return className; }
            set { className = value; }
        }

        public string Id
        {
            get { return id; }
            set { id = value; }
        }

    }
}
