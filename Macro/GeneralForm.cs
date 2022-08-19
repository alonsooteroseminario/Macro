using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Macro
{
    public partial class GeneralForm : Form
    {
        private string pathGeneral = @"C:\Users\"
                            + Environment.UserName
                            + @"\Box\PM Resources\Layout Team 2022\Repos Macros\FamiliesFiles\CAST IRON FITTINGS\";
        readonly int n = 26;
        private bool res;
        readonly int n2 = 26;
        private string num;
        private string type;

        bool active = true;
        readonly List<System.Windows.Forms.CheckBox> list_CheckBoxes = new List<CheckBox>();
        readonly List<System.Windows.Forms.CheckBox> list_CheckBoxesSize = new List<CheckBox>();
        public GeneralForm()
        {
            InitializeComponent();
        }
        private void Button3_Click(object sender, EventArgs e)
        {
            Hide();
            res = true;
            Exec();
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
                    foreach (var item in list_CheckBoxes)
                    {
                        if (item.Checked == true)
                        {
                            type = item.Text;
                        }
                    }
                    foreach (var item in list_CheckBoxesSize)
                    {
                        if (item.Checked == true)
                        {
                            num = item.Text;
                        }
                    }

                    string pathFile = pathGeneral + @"\" + type + @"\" + num + @"\";

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
        public void GetFoldersType()
        {
            this.groupBox1.Controls.Clear();
            int m = 0;

            var subFolders = Directory.GetDirectories(pathGeneral);

            foreach (var subFolder in subFolders)
            {
                System.Windows.Forms.CheckBox chBx = new System.Windows.Forms.CheckBox();
                this.groupBox1.Controls.Add(chBx);
                chBx.AutoSize = true;
                chBx.Location = new System.Drawing.Point(6, 23 + n * m);
                chBx.Size = new System.Drawing.Size(53, 20);
                chBx.UseVisualStyleBackColor = true;
                chBx.Name = "checkBox" + (m + 1).ToString();
                list_CheckBoxes.Add(chBx);

                m++;
                string[] splited = subFolder.Split('\\');
                chBx.Text = splited.Last();
            }
            foreach (var checkBox in list_CheckBoxes)
            {
                checkBox.CheckedChanged += new System.EventHandler(Custom_event_handler);
            }
        }
        private void Custom_event_handler(object sender, EventArgs e)
        {
            CheckBox cBox = sender as CheckBox;

            if (cBox.Checked == true)
            {
                foreach (var checkBox2 in list_CheckBoxes)
                {
                    if (cBox != checkBox2)
                    {
                        checkBox2.Checked = false;
                    }
                }
                GetFoldersSize();
            }
        }
        public string VerifyCheckBoxList()
        {
            string a = "";
            string b = "";
            foreach (var item in list_CheckBoxes)
            {
                if (item.Checked == true)
                {
                    a = item.Text;
                }
            }
            foreach (var item in list_CheckBoxesSize)
            {
                if (item.Checked == true)
                {
                    b = item.Text;
                }
            }

            string AB = "\\" + a + "\\" + b;
            return AB;
        }
        public void ChangeCheckBoxList()
        {
            try
            {
                var AB = VerifyCheckBoxList();

                var path = pathGeneral;

                var fileNames = Directory.GetFiles(path + AB);

                checkedListBox1.Items.Clear();
                foreach (var fileName in fileNames)
                {
                    string[] splited = fileName.Split('\\');
                    checkedListBox1.Items.Add(splited.Last()); // Full path
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void GetFoldersSize()
        {
            this.groupBox2.Controls.Clear();
            int m2 = 0;

            var subFolders = Directory.GetDirectories(pathGeneral);
            var firstSubFolder = subFolders.First();

            string[] splited_0 = firstSubFolder.Split('\\');
            string sizePath = pathGeneral + "\\" + splited_0.Last().ToString();
            var subFoldersSize = Directory.GetDirectories(sizePath);

            foreach (var subFolder in subFoldersSize)
            {
                System.Windows.Forms.CheckBox chBx = new System.Windows.Forms.CheckBox();
                this.groupBox2.Controls.Add(chBx);
                chBx.AutoSize = true;
                chBx.Location = new System.Drawing.Point(6, 23 + n2 * m2);
                chBx.Size = new System.Drawing.Size(42, 20);
                chBx.UseVisualStyleBackColor = true;
                chBx.Name = "checkBox" + (m2 + 14).ToString();
                list_CheckBoxesSize.Add(chBx);
                m2++;
                string[] splited = subFolder.Split('\\');
                chBx.Text = splited.Last();
            }
            foreach (var checkBox in list_CheckBoxesSize)
            {
                checkBox.CheckedChanged += new System.EventHandler(Custom_event_handlerSize);
            }
        }
        private void Custom_event_handlerSize(object sender, EventArgs e)
        {
            CheckBox cBox = sender as CheckBox;
            if (cBox.Checked == true)
            {
                foreach (var checkBox2 in list_CheckBoxesSize)
                {
                    if (cBox != checkBox2)
                    {
                        checkBox2.Checked = false;
                    }
                }
            }
            ChangeCheckBoxList();
        }
        private void BtnBrowseFile_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    pathGeneral = fbd.SelectedPath;
                    textBox2.Text = fbd.SelectedPath;
                    string[] splited = fbd.SelectedPath.Split('\\');
                    this.Text = splited.Last();
                    GetFoldersType();
                    GetFoldersSize();
                }
            }
        }
        private void GeneralForm_Activated(object sender, EventArgs e)
        {
            if (active)
            {
                string pathFile = @"C:\Users\"
                + Environment.UserName
                + @"\Box\PM Resources\Layout Team 2022\Repos Macros\FamiliesFiles\CAST IRON FITTINGS";

                string[] splited = pathFile.Split('\\');
                this.Text = splited.Last();
                
                pathGeneral = pathFile;
                
                textBox2.Text = pathFile;
                GetFoldersType();
                GetFoldersSize();
                active = false;
            }
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void CheckedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string itemText = checkedListBox1.Items[e.Index].ToString();

            var AB = VerifyCheckBoxList(); // "\\" + 22.5 + "\\" + 6;

            var path = pathGeneral;

            var totalFilePath = path + AB + "\\" + itemText;


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
                    OuterDB.ReadDwgFile(totalFilePath, System.IO.FileShare.Read, false, "");
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
    }
}
