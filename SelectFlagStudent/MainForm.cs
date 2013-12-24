using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using System.IO;

namespace SelectFlagStudent
{
    public partial class MainForm : BaseForm
    {
        BackgroundWorker BGW; //背景模式
        Workbook _WK;
        List<StudentObj> _FlagStudentList; //列印清單
        DateTime _StartTime, _EndTime; //日期區間

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            //取得今天日期
            DateTime today = DateTime.Today;
            //開始日期預設值
            dateTimeInput1.Value = today.AddDays(-7);
            //結束日期預設值
            dateTimeInput2.Value = today;
            BGW = new BackgroundWorker();
            BGW.WorkerReportsProgress = true;
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.ProgressChanged += new ProgressChangedEventHandler(BGW_ProgressChanged);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
        }

        private void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            LockForm(false); //解除表單元件封鎖
            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "留察記錄清單.xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _WK.Save(sd.FileName);
                    System.Diagnostics.Process.Start(sd.FileName);

                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    this.Enabled = true;
                    return;
                }
            }
        }

        private void BGW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage); //進度回報
        }

        private void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            GetStudents(); //取得資料
            _FlagStudentList.Sort(StudentSort); //列印清單排序
            Export(); //列印
            BGW.ReportProgress(100, "已完成 留察紀錄清單");
        }

        //清單排序方法
        private int StudentSort(StudentObj x, StudentObj y)
        {
            String xx = x.ClassName.PadLeft(20, '0');
            xx += x.SeatNo.PadLeft(3, '0');
            xx += x.Name.PadLeft(10, '0');

            String yy = y.ClassName.PadLeft(20, '0');
            yy += y.SeatNo.PadLeft(3, '0');
            yy += y.Name.PadLeft(10, '0');

            return xx.CompareTo(yy);
        }

        //取得資料
        private void GetStudents()
        {
            BGW.ReportProgress(10, "準備取得資料...");
            _FlagStudentList = new List<StudentObj>();
            //取得全部學生及班級
            List<StudentRecord> students = Student.SelectAll();
            List<ClassRecord> classes = Class.SelectAll();

            BGW.ReportProgress(20, "取得學生資料...");
            //建立學生字典快取
            Dictionary<String, StudentRecord> StudentCatch = new Dictionary<string, StudentRecord>();
            foreach (StudentRecord rec in students)
            {
                if (!StudentCatch.ContainsKey(rec.ID))
                {
                    StudentCatch.Add(rec.ID, rec);
                }
            }

            BGW.ReportProgress(30, "取得班級資料...");
            //建立班級字典快取
            Dictionary<String, String> ClassCatch = new Dictionary<string, string>();
            foreach (ClassRecord rec in classes)
            {
                if (!ClassCatch.ContainsKey(rec.ID))
                {
                    ClassCatch.Add(rec.ID, rec.Name);
                }
            }

            BGW.ReportProgress(40, "取得獎懲資料...");
            //取得獎懲紀錄
            List<DemeritRecord> records = K12.Data.Demerit.SelectByStudentIDs(StudentCatch.Keys.ToList());
            foreach (DemeritRecord record in records)
            {
                //該筆資料為留察
                if (record.MeritFlag == "2")
                {
                    string id = record.RefStudentID;
                    StudentRecord student = StudentCatch[id];
                    //該學生的狀態不是一般或延修就換下一筆
                    if (student.Status != StudentRecord.StudentStatus.一般 && student.Status != StudentRecord.StudentStatus.延修) continue;

                    DateTime time = record.OccurDate;
                    //該筆資料日期不在區間就換下一筆
                    if (time < _StartTime || time > _EndTime) continue;

                    string name = student.Name;
                    //座號不大於0就給空字串
                    string seatNo = (student.SeatNo > 0) ? student.SeatNo.ToString() : "";
                    string studentNo = student.StudentNumber;
                    string className;
                    //嘗試用學生對應班級代號取得班級名稱,遇到例外給空字串
                    try
                    {
                        className = ClassCatch[StudentCatch[id].RefClassID];
                    }
                    catch
                    {
                        className = "";
                    }

                    //建立該筆資料到列印清單中
                    _FlagStudentList.Add(new StudentObj(id, className, seatNo, studentNo, name, time));
                }
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (dateTimeInput1.ValueObject == null || dateTimeInput2.ValueObject == null)
            {
                MessageBox.Show("日期區間未輸入!");
                return;
            }

            //開始/結束日期區間
            _StartTime = dateTimeInput1.Value;
            _EndTime = dateTimeInput2.Value;
            //封鎖表單元件
            LockForm(true);
            //背景執行
            BGW.RunWorkerAsync();
        }

        //列印
        private void Export()
        {
            BGW.ReportProgress(50, "準備列印報表...");
            Workbook template = new Workbook();
            //開啟範本
            template.Open(new MemoryStream(Properties.Resources.留察紀錄清單));
            _WK = new Workbook();
            //複製範本
            _WK.Worksheets[0].Copy(template.Worksheets[0]);
            Worksheet ws = _WK.Worksheets[0];
            Cells cs = ws.Cells;

            //依據列印清單數量增加row
            Range oneRow = cs.CreateRange(2, 1, false);
            for (int i = 2; i < _FlagStudentList.Count + 2; i++)
            {
                cs.CreateRange(i, 1, false).Copy(oneRow);
            }

            //標題列印
            cs["A1"].PutValue(K12.Data.School.ChineseName + " 留察紀錄清單");
            cs["A2"].PutValue("班級");
            cs["B2"].PutValue("座號");
            cs["C2"].PutValue("學號");
            cs["D2"].PutValue("姓名");
            cs["E2"].PutValue("日期");

            BGW.ReportProgress(60, "開始列印報表...");
            int index = 2;
            float progress = 60;
            float rate = (float)(100-progress) / _FlagStudentList.Count; //進度百分比計算
            foreach(StudentObj student in _FlagStudentList)
            {
                progress += rate;
                BGW.ReportProgress((int)progress, "正在填入資料...");
                cs[index, 0].PutValue(student.ClassName);
                cs[index, 1].PutValue(student.SeatNo);
                cs[index, 2].PutValue(student.StudentNo);
                cs[index, 3].PutValue(student.Name);
                cs[index, 4].PutValue(student.Time.ToString("yyyy-MM-dd"));
                index++;
            }

        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //表單元件封鎖
        private void LockForm(bool b)
        {
            if(b)
            {
                dateTimeInput1.Enabled = false;
                dateTimeInput2.Enabled = false;
                buttonX1.Enabled = false;
            }
            else
            {
                dateTimeInput1.Enabled = true;
                dateTimeInput2.Enabled = true;
                buttonX1.Enabled = true;
            }
        }
    }
}
