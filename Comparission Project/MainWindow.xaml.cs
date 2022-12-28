using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Tekla.Structures.Model;

namespace Comparission_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<BeamComparission> List1 = new List<BeamComparission>();
        public List<BeamComparission> List2 = new List<BeamComparission>();
        string path = "comparission.txt";
        public MainWindow()
        {
            InitializeComponent();
        }
        public class BeamComparission
        {
            public string GUID { get; set; }
            public string StartPoint { get; set; }
            public string EndPoint { get; set; }
            public string Profile { get; set; }
            public string Class { get; set; }
            public BeamComparission(Beam beam)
            {
                this.StartPoint = beam.StartPoint.ToString();
                this.EndPoint = beam.EndPoint.ToString();
                this.Profile = beam.Profile.ProfileString;
                this.Class = beam.Class;
                this.GUID = beam.Identifier.GUID.ToString();
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(path))
            {
                File.Create(path);
            }
            Model model = new Model();

            var beams = model.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.BEAM);
            while (beams.MoveNext())
            {
                var current = beams.Current as Beam;
                if (current != null)
                {
                    List1.Add(new BeamComparission(current));
                }
            }
            System.Windows.MessageBox.Show("Close current model and open new model version");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Model model = new Model();

            var beams = model.GetModelObjectSelector().GetAllObjectsWithType(ModelObject.ModelObjectEnum.BEAM);
            while (beams.MoveNext())
            {
                var current = beams.Current as Beam;
                if (current != null)
                {
                    List2.Add(new BeamComparission(current));
                }
            }
            if (!File.Exists(path))
            {
                File.Create(path);
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            List<BeamComparission> list1 = List1;
            List<BeamComparission> list2 = List2;
            if (!File.Exists(path))
            {
                File.Create(path);
            }
            List<string> strings= new List<string>();

            strings.Add("model 1 Data");

            for (int i = 0; i < list1.Count(); i++)
            {
                string first = list1[i].GUID;
                string second = list1[i].StartPoint;
                string third = list1[i].EndPoint;
                string forth = list1[i].Profile;
                string fifth = list1[i].Class;
                string csvRow = string.Format("{0} \t{1}\t{2}\t{3}\t{4}", first, second, third, forth, fifth);
                
                strings.Add(csvRow);

            }
            strings.Add("\n");
            strings.Add("model 2 Data");

            for (int i = 0; i < list2.Count(); i++)
            {
                string first = list2[i].GUID;
                string second= list2[i].StartPoint;
                string third = list2[i].EndPoint;
                string forth = list2[i].Profile;
                string fifth = list2[i].Class;
                string row = string.Format("{0}\t{1}\t{2}\t{3}\t{4}", first, second, third, forth, fifth);
                strings.Add(row);
                
            }

            strings.Add("\n\n");
            strings.Add("Comparission for model 2");
            foreach (BeamComparission item2 in list2)
            {
                string message = CompareBothFile(list1, item2);
                strings.Add(message);
            }
            
            File.AppendAllLines(path, strings);
        }

        private string CompareBothFile(List<BeamComparission> list1, BeamComparission item2)
        {
            string guidError =string.Format("GUID {0} Not matched",item2.GUID);
            string spError = string.Format("Start Point {0} Not matched",item2.StartPoint);
            string epError = string.Format("End Point {0} Not matched",item2.EndPoint);
            string profileError = string.Format("Profile {0} Not matched",item2.Profile);
            string classError = string.Format("Class {0} Not matched",item2.Class);

            foreach (BeamComparission item in list1)
            {
                if (item2.GUID == item.GUID) guidError = string.Format("GUID {0} matched", item2.GUID);
                if (item2.StartPoint == item.StartPoint) spError = string.Format("StartPoint {0} matched", item2.StartPoint);
                if (item2.EndPoint == item.EndPoint) epError = string.Format("EndPoint {0} matched", item2.EndPoint);
                if (item2.Profile == item.Profile) profileError = string.Format("Profile {0} matched", item2.Profile);
                if (item2.Class == item.Class) classError = string.Format("Class {0} matched", item2.Class);
            }
            return guidError+"\n"+spError + "\n" +epError+"\n" +profileError + "\n" +classError + "\n";
            
        }
    }
}
