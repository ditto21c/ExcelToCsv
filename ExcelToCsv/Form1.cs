using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToCsv
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            LoadExcelFiles();

            Application.Exit();

        }

        void LoadExcelFiles()
        {
            CLoadExcel LoadExcel = new CLoadExcel();

            System.IO.DirectoryInfo DirectoryInfo = new System.IO.DirectoryInfo(System.Environment.CurrentDirectory);
            foreach (System.IO.FileInfo FileInfo in DirectoryInfo.GetFiles())
            {
                if (FileInfo.Extension == ".xlsx")
                {
                    LoadExcel.LoadExcel(FileInfo);
                }
            }
        }
    }
}