using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SelectFlagStudent
{
    class Permissions
    {
        public static string 留察紀錄清單 { get { return "SelectFlagStudent.DDA9F584-1303-4A04-851F-CBC093F24ADA"; } }

        public static bool 留察紀錄清單權限
        {
            get { return FISCA.Permission.UserAcl.Current[留察紀錄清單].Executable; }
        }
    }
}
