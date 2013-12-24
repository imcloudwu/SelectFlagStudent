using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SelectFlagStudent
{
    class Permissions
    {
        public static string 留查紀錄清單 { get { return "SelectFlagStudent.7B1D9E32-C8C7-4A5F-AEA8-25223BC7B980"; } }

        public static bool 留查紀錄清單權限
        {
            get { return FISCA.Permission.UserAcl.Current[留查紀錄清單].Executable; }
        }
    }
}
