using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Tekla.Structures.Model;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using System.Windows.Input;
using System;
using System.Data.Common;
using Tekla.Structures.Solid;

namespace Comparission_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public List<BeamComparison> List1 = new List<BeamComparison>();
        public List<BeamComparison> List2 = new List<BeamComparison>();
        string path = "Comparison.txt";
        string xlPath = "O.xlsx";
        public MainWindow()
        {
            InitializeComponent();
        }
        public class BeamComparison
        {
            public string GUID { get; set; }
            public string StartPoint { get; set; }
            public string EndPoint { get; set; }
            public string Profile { get; set; }
            public string Class { get; set; }
            public BeamComparison(Beam beam)
            {
                this.StartPoint = beam.StartPoint.ToString();
                this.EndPoint = beam.EndPoint.ToString();
                this.Profile = beam.Profile.ProfileString;
                this.Class = beam.Class;
                this.GUID = beam.Identifier.GUID.ToString();
            }
        }

        public List<string> CreateRowFromData(BeamComparison beamComparission)
        {
            List<string> row = new List<string>();
            row.Add(beamComparission.GUID);
            row.Add(beamComparission.StartPoint);
            row.Add(beamComparission.EndPoint);
            row.Add(beamComparission.Profile);
            row.Add(beamComparission.Class);
            return row;
        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Tekla.Structures.Model.Model model = new Tekla.Structures.Model.Model();
            if(!model.GetConnectionStatus()) System.Windows.MessageBox.Show("Unable to connect to Tekla");
            var beams = model.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.BEAM);
            while (beams.MoveNext())
            {
                var current = beams.Current as Beam;
                if (current != null)
                {
                    List1.Add(new BeamComparison(current));
                }
            }
            System.Windows.MessageBox.Show("Close current model and open new model version");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Tekla.Structures.Model.Model model = new Tekla.Structures.Model.Model();
            var beams = model.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.BEAM);
            while (beams.MoveNext())
            {
                var current = beams.Current as Beam;
                if (current != null)
                {
                    List2.Add(new BeamComparison(current));
                }
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
           
            xlPath = System.IO.Path.GetDirectoryName(strExeFilePath)+"\\Report.xlsx";
            
            List<BeamComparison> list1 = List1;
            List<BeamComparison> list2 = List2;
            
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(xlPath);
            _Worksheet Initial = xlWorkbook.Sheets[1];
            _Worksheet Final = xlWorkbook.Sheets[2];
            _Worksheet Comparission = xlWorkbook.Sheets[3];
            List<string> title = new List<string>() { "GUID", "StartPoint", "EndPoint", "Profile", "Class" };
            List<string> titleFinal = new List<string>() { "GUID_First Model", "StartPoint_First Model", "EndPoint_First Model", "Profile_First Model", "Class_First Model", "GUID_Second Model", "StartPoint_Second Model", "EndPoint_Second Model", "Profile_Second Model", "Class_Second Model", "Check", "Remark" };

            PrintRowData(Initial, 1, title,1);
            PrintRowData(Final, 1, title, 1);
            PrintRowData(Comparission, 1, titleFinal, 1);

            PrintSheet(Initial, list1,1);
            PrintSheet(Final, list2,1);

            PrintComparisonSheet(Comparission, list1, list2);

            xlApp.Visible = true;
            xlApp.UserControl = true;
            xlWorkbook.Save();
        }

        private void PrintComparisonSheet(_Worksheet comparission, List<BeamComparison> list1, List<BeamComparison> list2)
        {
            PrintSheet(comparission, list1, 1);
            int row = 2;
            for (int i = 0; i < list2.Count; i++)
            {
                bool isMatched=false;
                List<string> toPrint = CreateRowFromData(list2[i]);
                for (int j = 0; j < list1.Count; j++)
                {
                    if (IsSameCoordinate(list1[j], list2[i]))
                    {
                        isMatched = true;
                        string remark = "";
                        if (list1[j].Profile != list2[i].Profile)
                        {
                            remark = remark + "Profile Changed. ";
                        }
                        if (list1[j].GUID != list2[i].GUID)
                        {
                            remark = remark + "GUID Changed. ";
                        }
                        if (list1[j].Class != list2[i].Class)
                        {
                            remark = remark + "Class Changed. ";
                        }
                        if (remark == "")
                        {
                            toPrint.Add("Checked-OK");
                        }
                        else
                        {
                            toPrint.Add("Checked-Failed");
                        }
                        toPrint.Add(remark);
                    }
                }
                if (!isMatched) 
                {
                    toPrint.Add("Checked-Failed");
                    toPrint.Add("New Object Detected");
                }
                PrintRowData(comparission, row, toPrint, 6);
                row++;
            }
        }

        private void PrintSheet(_Worksheet sheet, List<BeamComparison> list,int column)
        {
            int rowNumber = 2;
            for (int i = 0; i < list.Count; i++)
            {
                PrintRowData(sheet, rowNumber, CreateRowFromData(list[i]),column);
                rowNumber++;
            }
        }

        private void PrintRowData(_Worksheet initial, int rowNumber, List<string> list, int column)
        {
            for (int i = 0; i < list.Count; i++)
            {
                initial.Cells[rowNumber, column] = list[i];
                column++;
            }
        }
        public bool IsSameCoordinate(BeamComparison beam1, BeamComparison beam2)
        {
            var result = false;
            if (beam1.StartPoint == beam2.StartPoint && beam1.EndPoint == beam2.EndPoint)
            {
                result = true;
            }
            return result;
        }
    }
}
