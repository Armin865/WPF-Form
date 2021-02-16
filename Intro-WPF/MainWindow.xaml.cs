using Ganss.Excel;
using Microsoft.PowerBI.Api.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
using excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;




//Armin Prutina
//WPF form that lets users save first and last names into either Text form, Excel form, or SQL server



namespace Intro_WPF
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Path = "";

        }
        //Sets the string for First and Last name
        string FirstName;
        String LastName;
        public string Path { get; private set; }



        //Setting up SQL connection to allow program to enter 
        SqlConnection con = new SqlConnection(@"Data Source=localhost\SQLEXPRESS;Initial Catalog=NABA;Integrated Security=True;");
        SqlCommand cmd;





        //Writes into the Textfile using button
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            StreamWriter sw = new StreamWriter("c:/users/adipr/source/repos/Intro-WPF/ReadandWrite.txt");
            sw.WriteLine(FirstName + ' ' + LastName);
            sw.Close();
        }



        //setting the Firstname to textbox 1
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FirstName = txtbox1.Text;
        }


        //setting the lastname to textbox 2
        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            LastName = txtbox2.Text;
        }



        //Setting 2nd button to write to Excel file
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var students = new List<NAME>
            { //getting the first and last name input from above
                new NAME{FirstName=FirstName},
                new NAME{LastName=LastName}
            };
            //using ExcelMapper and creating a filename for XLSX
            ExcelMapper mapper = new ExcelMapper();
            var newFile = ("c:/users/adipr/source/repos/Intro-WPF/ReadandWrite.xlsx");
            mapper.Save(newFile, students, "SheerName", true);
           
        }




        //setting up the public class for NAME
        public class NAME
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
        }


        //Setting 3rd button to write to SQL Server Database 
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            con.Open();
            cmd = new SqlCommand("Insert into Person values('" + txtbox1.Text + "','" + txtbox2.Text + "')", con);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Data has been saved to NABA Database! ");
            con.Close();
        }
    }

    
    }


