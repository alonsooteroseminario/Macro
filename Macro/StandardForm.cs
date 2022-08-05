using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Macro
{
    public partial class StandardForm : Form
    {
        private string num;
        private string type;
        private bool res;
        private string pathGeneral;

        public StandardForm(string pathGeneral)
        {
            this.pathGeneral = pathGeneral;
        }

        public StandardForm()
        {
            InitializeComponent();
        }
        private void Button1_Click_1(object sender, EventArgs e)
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
                        type = "TEE-WYE";
                    }
                    if (checkBox2.Checked == true)
                    {
                        type = "90";
                    }
                    if (checkBox5.Checked == true)
                    {
                        type = "BOSTON";
                    }
                    if (checkBox4.Checked == true)
                    {
                        type = "22.5";
                    }
                    if (checkBox3.Checked == true)
                    {
                        type = "45";
                    }
                    if (checkBox6.Checked == true)
                    {
                        type = "WYE";
                    }
                    if (checkBox7.Checked == true)
                    {
                        type = "P-TRAP";
                    }
                    if (checkBox8.Checked == true)
                    {
                        type = "Laundry";
                    }
                    if (checkBox9.Checked == true)
                    {
                        type = "OFFSETS";
                    }
                    if (checkBox10.Checked == true)
                    {
                        type = "REDUCER";
                    }
                    if (checkBox11.Checked == true)
                    {
                        type = "Kitchen Sink";
                    }
                    if (checkBox12.Checked == true)
                    {
                        type = "BTG";
                    }
                    if (checkBox13.Checked == true)
                    {
                        type = "BV_VS";
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
                    if (checkBox23.Checked == true)
                    {
                        num = "6";
                    }
                    if (checkBox24.Checked == true)
                    {
                        num = "8";
                    }

                    string pathFile = @"C:\\Users\\" +
                        Environment.UserName +
                        @"\\Box\\PM Resources\\Layout Team 2022\\Repos Macros\\FamiliesFiles\\CAST IRON FITTINGS\\" +
                         type + @"\\" +
                         num + @"\\";
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (checkedListBox1.CheckedItems.Contains(checkedListBox1.Items[i]))
                        {
                            var spl = checkedListBox1.Items[i].ToString().Split('\\');
                            string letter = spl.Last();
                            MyCommands myCommands = new MyCommands();
                            myCommands.CASTFITTINGS(pathFile, ppr, letter);
                        }
                    }

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
            }
        }
        private void CheckBox18_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox18;
            if (cBox.Checked == true)
            {
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
            }
        }
        private void CheckBox15_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox15;
            if (cBox.Checked == true)
            {
                checkBox18.Checked = false;
                checkBox19.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
            }
        }
        private void CheckBox19_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox19;
            if (cBox.Checked == true)
            {
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
            }
        }
        private void CheckBox23_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox23;
            if (cBox.Checked == true)
            {
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox24.Checked = false;
            }
        }
        private void CheckBox24_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cBox = checkBox24;
            if (cBox.Checked == true)
            {
                checkBox18.Checked = false;
                checkBox15.Checked = false;
                checkBox19.Checked = false;
                checkBox23.Checked = false;
            }
        }
        private void BtnBrowseFolder_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    pathGeneral = fbd.SelectedPath;
                    textBox1.Text = fbd.SelectedPath;
                }
            }
        }

        public void ChangeCheckBoxList()
        {
            try
            {
                string a = "";
                string b = "";
                if (checkBox4.Checked == true)
                {
                    a = "22.5";
                }
                if (checkBox3.Checked == true)
                {
                    a = "45";
                }
                if (checkBox18.Checked == true)
                {
                    b = "2";
                }
                if (checkBox15.Checked == true)
                {
                    b = "3";
                }
                if (checkBox19.Checked == true)
                {
                    b = "4";
                }
                if (checkBox23.Checked == true)
                {
                    b = "6";
                }
                if (checkBox24.Checked == true)
                {
                    b = "8";
                }

                var fileNames = Directory.GetFiles("C:\\Users\\" +
                    Environment.UserName +
                    "\\Box\\" + "PM Resources\\Layout Team 2022\\Repos Macros\\FamiliesFiles\\" + "CAST IRON FITTINGS\\" +
                    a +
                    "\\" +
                    b);


                checkedListBox1.Items.Clear();
                foreach (var fileName in fileNames)
                {
                    //char fileNameToShow = fileName[fileName.Length-5];
                    checkedListBox1.Items.Add(fileName); // Full path
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void CheckedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string itemText = checkedListBox1.Items[e.Index].ToString();

            for (int ix = 0; ix < checkedListBox1.Items.Count; ++ix)
            {
                if (ix != e.Index) checkedListBox1.SetItemChecked(ix, false);
            }
            if (e.NewValue == CheckState.Checked)
            {
                //MessageBox.Show(itemText + " checked");
                _ = new ObjectIdCollection();
                using (Database OuterDB = new Database())
                {
                    OuterDB.ReadDwgFile(itemText, System.IO.FileShare.Read, false, "");
                    using (Transaction tr = OuterDB.TransactionManager.StartTransaction())
                    {
                        BlockTable bt;
                        bt = (BlockTable)tr.GetObject(OuterDB.BlockTableId
                                                       , OpenMode.ForRead);

                        BlockTableRecord blk = (BlockTableRecord)tr.GetObject(bt["*Model_Space"], OpenMode.ForRead);

                        var imgsrc = Autodesk.AutoCAD.Windows.Data.CMLContentSearchPreviews.GetBlockTRThumbnail(blk);
                        var bmp = ImageSourceToGDI(imgsrc as System.Windows.Media.Imaging.BitmapSource);

                        //var image = new Form1(bmp as System.Drawing.Bitmap);
                        //image.ShowDialog();

                        pictureBox2.BackgroundImageLayout = ImageLayout.Stretch;
                        pictureBox2.BackgroundImage = bmp;

                        tr.Commit();
                    }
                }
            }
        }

        private static System.Drawing.Image ImageSourceToGDI(System.Windows.Media.Imaging.BitmapSource src)
        {
            var ms = new MemoryStream();
            var encoder =
              new System.Windows.Media.Imaging.BmpBitmapEncoder();
            encoder.Frames.Add(
              System.Windows.Media.Imaging.BitmapFrame.Create(src)
            );
            encoder.Save(ms);
            ms.Flush();
            return System.Drawing.Image.FromStream(ms);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            ChangeCheckBoxList();
        }

        public override bool Equals(object obj)
        {
            return obj is StandardForm form &&
                   pathGeneral == form.pathGeneral;
        }
    }
}
