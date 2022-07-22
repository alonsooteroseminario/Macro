﻿using Autodesk.AutoCAD.ApplicationServices;
using System;
using System.Windows.Forms;

namespace Macro
{
    public partial class LayersForm : Form
    {
        private Document doc;
        private string pathGeneral;
        public LayersForm()
        {
            InitializeComponent();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            MyCommands myCommands = new MyCommands();
            myCommands.LAYERS(pathGeneral, doc);
        }
        private void btnBrowseFile_Click(object sender, EventArgs e)
        {
            using (var fbd = new OpenFileDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    pathGeneral = fbd.FileName;
                    textBox2.Text = fbd.FileName;
                }
            }
        }
        private void LayersForm_Activated(object sender, EventArgs e)
        {
            string pathFile = @"C:\Users\"
                + Environment.UserName
                + @"\Box\PM Resources\Layout Team 2022\LAYERS_TO_TURN_OFF.xlsx";

            pathGeneral = pathFile;
            textBox2.Text = pathFile;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
