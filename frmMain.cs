using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Random_Dates
{
    public partial class frmMain : Form
    {
        private Random gen = new Random();
        List<UserSchedule> schedules = new List<UserSchedule>();
        List<UserInput> inputs = new List<UserInput>();
        //UserSchedule schedule;
        DataTable dt = new DataTable();
        public frmMain()
        {
            InitializeComponent();
        }

        //the default would be 6 dates that fall between
        //the date/time range (time is important)
        //When I enter a new entry the date/times cannot
        //overlap with the previous ones that where entered
        //they have to be at least 30 minutes apart
        //If the wellness flag is selected I need two more dates
        //generated that are 15 minutes later than any 2 of the
        //previously generated dates
        DateTime RandomDay(DateTime start, DateTime stop)
        {
            int range = (int)(stop - start).TotalMinutes;
            return start.AddMinutes(gen.Next(0, range));
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            dtpFrom.Value = DateTime.Now.AddHours(1);
            dtpTo.Value = DateTime.Now.AddHours(24);

            var c = dt.Columns.Add("Date/Time");
            c.ReadOnly = true;
            dgv1.DataSource = dt;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtName.Text))
            {
                MessageBox.Show("User's name is required!"); return;
            }

            inputs.Clear();
            inputs.Add(new UserInput(
                1, txtName.Text, 
                dtpFrom.Value, dtpTo.Value, 
                ckbWellness.Checked, ckbAdult.Checked));

            pb.Maximum = GetDateCount(ckbWellness.Checked, ckbAdult.Checked);
            pb.Value = 1;

            btnGenerate.Enabled = false;
            btnCancel.Enabled = true;
            backgroundWorker1.RunWorkerAsync();
        }

        private void DisplaySchedules()
        {
            dt.Rows.Clear();
            dt.Columns.Clear();
            var c = dt.Columns.Add("Date/Time");
            c.ReadOnly = true;
            foreach (UserSchedule userSchedule in schedules) {
                int cindex = schedules.IndexOf(userSchedule) + 1;

                if(cindex == dt.Columns.Count) {
                    dt.Columns.Add(userSchedule.Name);
                }

                foreach (var item in userSchedule.dates)
                {
                    int index = -1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                        if (dt.Rows[i][0].ToString() ==
                            item.ToShortDateString())
                        { index = i; break; }

                    if (index == -1) {
                        index = dt.Rows.Count;
                        for (int i = dt.Rows.Count - 1; i >= 0; i--)
                        {
                            if (Convert.ToDateTime(dt.Rows[i][0]) > 
                                Convert.ToDateTime(item.ToShortDateString())) index = i;
                        }
                        var dr = dt.NewRow();
                        dr[0] = item.ToShortDateString();
                        dt.Rows.InsertAt(dr, index);
                    }

                    string str = dt.Rows[index][cindex].ToString();
                    if (str.Length > 0) str += ", ";
                    str += item.ToShortTimeString();
                    dt.Rows[index][cindex] = str;
                }
            }
        }

        private bool ValidateDate(DateTime date, List<DateTime> dates, bool use15mins)
        {
            int n = use15mins ? 15 : 30;
                //(x <= dt && (dt - x.AddMinutes(n)).TotalMinutes >= n) ||
            return dates.Count(x =>
                (x <= date && (x.AddMinutes(n) >= date)) ||
                (x >= date && (x.AddMinutes(-n) >= date))) < 1;
        }
        private int GetDateCount(bool wellness, bool adult)
        {
            int i = 5;
            if (!adult) i = 3;
            if (wellness) i += 2;
            return i;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<DateTime> dates = new List<DateTime>();
            foreach (var item in schedules)
                dates.AddRange(item.dates);

            foreach (var input in inputs)
            {
                var schedule = schedules.FirstOrDefault(x => x.Name == input.Name);
                if (schedule == null) {
                    schedule = new UserSchedule(schedules.Count + 1,
                        input.Name, new List<DateTime>());
                    schedules.Add(schedule);
                }
                int dateCount = GetDateCount(input.isWellness, input.isAdult);
                for (int i = 0; i < dateCount; i++)
                {
                    DateTime date = RandomDay(input.fromDate, input.toDate);
                    while (!ValidateDate(date, dates, input.isWellness && i >= dateCount - 2))
                    {
                        date = RandomDay(input.fromDate, input.toDate);
                    }
                    schedule.dates.Add(date);
                    backgroundWorker1.ReportProgress(i);
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DisplaySchedules();
            pb.Value = pb.Maximum;
            btnGenerate.Enabled = true;
            btnImport.Enabled = true;
            btnExport.Enabled = true;
            btnCancel.Enabled = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                if(backgroundWorker1.IsBusy) backgroundWorker1.CancelAsync();
            }
            catch (Exception)
            {
            }
            btnGenerate.Enabled = true;
            btnImport.Enabled = true;
            btnExport.Enabled = true;
            btnCancel.Enabled = false;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pb.Value = e.ProgressPercentage;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (dt.Rows.Count > 0)
                {
                    SaveFileDialog sfd = new SaveFileDialog()
                    {
                        Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
                    };
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            if (sfd.FileName.EndsWith(".csv"))
                            {
                                string str = "";
                                for (int i = 0; i < dt.Columns.Count; i++)
                                    str += (str.Length == 0 ? "" : ",") + dt.Columns[i].ColumnName;
                                foreach (DataRow r in dt.Rows)
                                {
                                    str += Environment.NewLine;
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                        str += (i == 0 ? "" : ",") + 
                                            (!r[i].ToString().Contains(",") ? 
                                            r[i] : "\"" + r[i] + "\"");
                                }
                                System.IO.File.WriteAllText(sfd.FileName, str);
                            }
                            else
                                new ExcelHandler().ExportTable(dt, sfd.FileName);
                            MessageBox.Show("Record Exported!");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    //Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
                    Filter = "Excel Files (*.xlsx)|*.xlsx"
                };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var dtImport = 
                            new ExcelHandler().ImportTable(ofd.FileName);
                        if (MessageBox.Show(
                            "Record Imported!\n\rDo you want to process this input?",
                            "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            //dgv1.DataSource= dtImport;
                            string str1 = dtImport.Rows[0][1].ToString();
                            string str2 = dtImport.Rows[1][1].ToString();

                            str1 = str1.Split(' ')[0];
                            str2 = str2.Split(' ')[0];

                            inputs.Clear();
                            int Maximum = 0;
                            for (int i = 3; i < dtImport.Rows.Count; i++)
                            {
                                inputs.Add(new UserInput(
                                    i - 2, dtImport.Rows[i][0].ToString(),
                                    Convert.ToDateTime(str1 + " " + dtImport.Rows[i][1].ToString()), 
                                    Convert.ToDateTime(str2 + " " + dtImport.Rows[i][2].ToString()),
                                    dtImport.Rows[i][3].ToString().ToLower() == "yes", 
                                    dtImport.Rows[i][4].ToString().ToLower() == "yes"
                                ));
                                Maximum += GetDateCount(inputs.Last().isWellness, inputs.Last().isAdult);
                            }

                            pb.Maximum = Maximum * inputs.Count;
                            pb.Value = 1;

                            btnGenerate.Enabled = false;
                            btnImport.Enabled = false;
                            btnExport.Enabled = false;
                            btnCancel.Enabled = true;
                            backgroundWorker1.RunWorkerAsync();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
    public class UserInput
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime fromDate { get; set; }
        public DateTime toDate { get; set; }
        public bool isWellness { get; set; }
        public bool isAdult { get; set; }
                   
        public UserInput(
            int id, string name, 
            DateTime fromDate, DateTime toDate, 
            bool isWellness, bool isAdult)
        {
            this.Id = id;
            this.Name = name;
            this.fromDate = fromDate;
            this.toDate = toDate;
        }
    }

    public class UserSchedule
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public List<DateTime> dates { get; set; }

        public UserSchedule(int id, string name, List<DateTime> dates)
        {
            this.Id = id;
            this.Name = name;
            this.dates = dates;
        }
    }
}
