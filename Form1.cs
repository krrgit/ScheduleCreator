using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab


namespace ScheduleCreator
{
    public partial class ScheduleCreator : Form
    {

        DataSet availDataSet;
        int[] countEmpAM = new int[7];
        int[] countEmpPM = new int[7];
        ListBox[] listBoxesAM = new ListBox[7];
        ListBox[] listBoxesPM = new ListBox[7];

        public ScheduleCreator()
        {
            InitializeComponent();
            SetFirstDayOfWeek();
            UpdateHourHeader();
            InitializeEmployeeCounts();
            InitializeEmployeeLists();
        }

        // Sets First Day of Week to next Monday
        private void SetFirstDayOfWeek() 
        {
            dateTimePicker1.Value = DateTime.Today.AddDays(((int)DateTime.Today.DayOfWeek - (int)DayOfWeek.Monday) + 7); ;
        }

        private void InitializeEmployeeCounts() {
            countEmpAM[0] = Convert.ToInt32(countMonAM.Text);
            countEmpAM[1] = Convert.ToInt32(countTueAM.Text);
            countEmpAM[2] = Convert.ToInt32(countWedAM.Text);
            countEmpAM[3] = Convert.ToInt32(countThrAM.Text);
            countEmpAM[4] = Convert.ToInt32(countFriAM.Text);
            countEmpAM[5] = Convert.ToInt32(countSatAM.Text);
            countEmpAM[6] = Convert.ToInt32(countSunAM.Text);

            countEmpPM[0] = Convert.ToInt32(countMonPM.Text);
            countEmpPM[1] = Convert.ToInt32(countTuePM.Text);
            countEmpPM[2] = Convert.ToInt32(countWedPM.Text);
            countEmpPM[3] = Convert.ToInt32(countThrPM.Text);
            countEmpPM[4] = Convert.ToInt32(countFriPM.Text);
            countEmpPM[5] = Convert.ToInt32(countSatPM.Text);
            countEmpPM[6] = Convert.ToInt32(countSunPM.Text);
        }

        private void InitializeEmployeeLists()
        {
            listBoxesAM[0] = listMonAM;
            listBoxesAM[1] = listTueAM;
            listBoxesAM[2] = listWedAM;
            listBoxesAM[3] = listThrAM;
            listBoxesAM[4] = listFriAM;
            listBoxesAM[5] = listSatAM;
            listBoxesAM[6] = listSunAM;

            listBoxesPM[0] = listMonPM;
            listBoxesPM[1] = listTuePM;
            listBoxesPM[2] = listWedPM;
            listBoxesPM[3] = listThrPM;
            listBoxesPM[4] = listFriPM;
            listBoxesPM[5] = listSatPM;
            listBoxesPM[6] = listSunPM;
        }

        private void OpenAvailabilityFile_Click(object sender, EventArgs e)
        {
            OleDbConnection conn;
            OleDbDataAdapter dta;
            string excel;
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|Xls Files (*.xls)|*.xls|All Files(*.*)|*.*";
            dialog.Title = "Import Availability File";

            try
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fi = new FileInfo(dialog.FileName);
                    string fileName = dialog.FileName;
                    excel = fi.FullName;
                    conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excel +";Extended Properties=Excel 12.0;");
                    dta = new OleDbDataAdapter("select * from [Sheet1$]", conn);
                    availDataSet = new DataSet();
                    dta.Fill(availDataSet, "[Sheet1$]");
                    availDataGrid.DataSource = availDataSet;
                    availDataGrid.DataMember = "[Sheet1$]";
                    conn.Close();
                    FormatAvailibility();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FormatAvailibility() 
        {
            availDataGrid.Columns[0].Width = 120;
            for (int i = 1; i < availDataGrid.Columns.Count; i++)
            {
                availDataGrid.Columns[i].Width = 70;
            }
        }

        private void GenerateSchedule()
        {
            for (int i = 1; i <=7; ++i)
            {
                SortNames(i);
            }
        }

        private void SortNames(int day) 
        {
            var dayIndex = day - 1;
            int dayAMCount = countEmpAM[dayIndex];
            int dayPMCount = countEmpPM[dayIndex];

            var count = new int[2];
            var availText = "";
            for (int i = 1; i < availDataGrid.RowCount - 1; ++i)
            {
                availText = availDataGrid.Rows[i].Cells[day].Value.ToString();
                if (availText.Contains("Morning") && count[0] < dayAMCount)
                {
                    listBoxesAM[dayIndex].Items.Add(availDataGrid.Rows[i].Cells[0].Value);
                    ++count[0];
                } 
                else if (availText.Contains("Night") && count[1] < dayPMCount)
                {
                    listBoxesPM[dayIndex].Items.Add(availDataGrid.Rows[i].Cells[0].Value);
                    ++count[1];
                }
            }

            if (count[0] == dayAMCount || count[1] == dayPMCount)
            {
                return;
            }

            for (int i = 1; i < availDataGrid.RowCount - 1; ++i)
            {
                availText = availDataGrid.Rows[i].Cells[day].Value.ToString();
                if (availText.Contains("Any"))
                {
                    if (count[0] < dayAMCount)
                    {
                        listBoxesAM[dayIndex].Items.Add(availDataGrid.Rows[i].Cells[0].Value);
                        ++count[0];
                    }
                    else if (count[1] < dayPMCount)
                    {
                        listBoxesPM[dayIndex].Items.Add(availDataGrid.Rows[i].Cells[0].Value);
                        ++count[1];
                    }
                    else 
                    {
                        break;
                    }

                }
                 
            }
            return;
        }



        private void GenerateBtn_Click(object sender, EventArgs e)
        {
            GenerateSchedule();
            ScheduleTab.SelectedTab = tabDay;
            MessageBox.Show("Schedule Generated.");
        }

        private void StartTimeUpdate_Click(object sender, EventArgs e)
        {
            UpdateHourHeader();
        }

        private void UpdateHourHeader()
        {
            int startIndex = cBoxStartTime.SelectedIndex - (cBoxStartTime.SelectedItem.ToString().Contains("30") ? 1 : 0);
            string label = "";
            string tempHour = "";

            for (int i = startIndex; i < cBoxStartTime.Items.Count; i += 2)
            {
                if (!cBoxStartTime.Items[i].ToString().Contains("30"))
                {
                    tempHour = cBoxStartTime.Items[i].ToString().Split(':')[0];
                    tempHour = tempHour.Length == 1 ? tempHour.Insert(0, " ") : tempHour;
                    label += tempHour + "  |  ";
                }
            }
            labelHourHeader.Text = label;
        }
    }
}
