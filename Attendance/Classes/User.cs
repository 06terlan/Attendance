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
        public SortedDictionary<DateTime, List<string>> inOutType;

        public User(int id)
        {
            this.id = id;
            //allInOuts = new List<DateTime>();
            inOutType = new SortedDictionary<DateTime, List<string>>();
        }

        public int entered(DateTime dt)
        {
            key(dt);
            inOutType[dt].Add("in");

            return ++allIns;
        }

        public int exited(DateTime dt)
        {
            key(dt);
            inOutType[dt].Add("out");

            return ++allOuts;
        }

        private bool key(DateTime dt)
        {

            if (!inOutType.ContainsKey(dt))
            {
                inOutType[dt] = new List<string>();
            }

            return true;
        }
    }
}
