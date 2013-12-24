using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SelectFlagStudent
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學務作業", "資料統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            item1["報表"]["留查紀錄清單"].Enable = Permissions.留查紀錄清單權限;
            item1["報表"]["留查紀錄清單"].Click += delegate
            {
                MainForm f = new MainForm();
                f.ShowDialog();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["學務作業"]["資料統計"];
            permission.Add(new RibbonFeature(Permissions.留查紀錄清單, "留查紀錄清單"));
        }
    }
}

