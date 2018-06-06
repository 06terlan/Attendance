using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Attendance.Classes
{
    class User
    {
        public int allOuts = 0;
        public int allIns = 0;
        public int id;
        //public List<DateTime> allInOuts;
        public SortedDictionary<DateTime, string> inOutType;

        public User(int id)
        {
            this.id = id;
            //allInOuts = new List<DateTime>();
            inOutType = new SortedDictionary<DateTime, string>();
        }

        public int entered(DateTime dt)
        {
            //allInOuts.Add(dt);
            inOutType[dt] = "in";

            return ++allIns;
        }

        public int exited(DateTime dt)
        {
            //allInOuts.Add(dt);
            inOutType[dt] = "out";

            return ++allOuts;
        }
    }
}
