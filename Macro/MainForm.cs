using Autodesk.AutoCAD.EditorInput;
using System;
using System.Windows.Forms;

namespace Macro
{
    public partial class MainForm : Form
    {
        private string num;
        private string discipline;
        private bool res;
        private readonly string pathGeneral;
        public MainForm()
        {
            InitializeComponent();
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            Hide();
            res = true;
            Exec();
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void Exec()
        {
            while (res)
            {
                Editor edt = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                PromptPointOptions ppo = new PromptPointOptions("Select the point where to place");
                PromptPointResult ppr = edt.GetPoint(ppo);
                if (ppr.Status == PromptStatus.Cancel)
                {
                    res = false;
                    Show();
                }
                if (ppr.Status == PromptStatus.OK)
                {
                    if (checkBox1.Checked == true)
                    {
                        discipline = "SANITARY";
                    }
                    if (checkBox2.Checked == true)
                    {
                        discipline = "STORM";
                    }
                    if (checkBox5.Checked == true)
                    {
                        discipline = "DCW";
                    }
                    if (checkBox4.Checked == true)
                    {
                        discipline = "GAS";
                    }
                    if (checkBox3.Checked == true)
                    {
                        discipline = "VENT";
                    }
                    if (checkBox6.Checked == true)
                    {
                        discipline = "HPS";
                    }
                    if (checkBox7.Checked == true)
                    {
                        discipline = "HPR";
                    }
                    if (checkBox8.Checked == true)
                    {
                        discipline = "HWS";
                    }
                    if (checkBox9.Checked == true)
                    {
                        discipline = "HWR";
                    }
                    if (checkBox10.Checked == true)
                    {
                        discipline = "CWS";
                    }
                    if (checkBox11.Checked == true)
                    {
                        discipline = "CWR";
                    }
                    if (checkBox12.Checked == true)
                    {
                        discipline = "DHW";
                    }
                    if (checkBox13.Checked == true)
                    {
                        discipline = "DHWR";
                    }
                    if (checkBox14.Checked == true)
                    {
                        discipline = "COND";
                    }
                    if (checkBox16.Checked == true)
                    {
                        num = "1";
                    }
                    if (checkBox17.Checked == true)
                    {
                        num = "1.5";
                    }
                    if (checkBox18.Checked == true)
                    {
                        num = "2";
                    }
                    if (checkBox15.Checked == true)
                    {
                        num = "3";
                    }
                    if (checkBox19.Checked == true)
                    {
                        num = "4";
                    }
                    if (checkBox21.Checked == true)
                    {
                        num = "5";
                    }
                    if (checkBox23.Checked == true)
                    {
                        num = "6";
                    }
                    if (checkBox24.Checked == true)
                    {
                        num = "8";
                    }
                    if (checkBox22.Checked == true)
                    {
                        num = "10";
                    }
                    if (checkBox20.Checked == true)
                    {
                        num = "12";
                    }
                    string pathFile = pathGeneral + @"\"
                        + discipline + @"\"
                        + num + @"\";
                    MyCommands myCommands = new MyCommands();
                    myCommands.CAN(num, discipline, ppr, pathFile);
                }
            }
        }
        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox1;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox2;
            if (cBox.Checked == true)
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox3;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox1.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox5_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox5;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox1.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox4;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox1.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox6_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox6;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox1.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox7_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox7;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox1.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox8_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox8;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox1.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox9_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox9;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox1.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox10_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox10;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox1.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox11_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox11;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox1.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox12_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox12;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox1.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox13_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox13;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox1.Checked = false;
                checkBox14.Checked = false;
            }
        }
        private void CheckBox14_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox14;
            if (cBox.Checked == true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox1.Checked = false;
            }
        }
        private void CheckBox16_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox16;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox22.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox17_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox17;
            if (cBox.Checked == true)
            {
                checkBox16.Checked = false;
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox22.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox18_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox18;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox16.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox22.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox15_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox15;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox16.Checked = false;
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox22.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox19_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox19;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox16.Checked = false;
                checkBox21.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox22.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox21_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox21;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox16.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox22.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox23_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox23;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox16.Checked = false;
                checkBox24.Checked = false;
                checkBox22.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox24_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox24;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox23.Checked = false;
                checkBox16.Checked = false;
                checkBox22.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox22_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox22;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox16.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void CheckBox20_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox20;
            if (cBox.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox21.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox22.Checked = false;
                checkBox16.Checked = false;
            }
        }
    }
}
