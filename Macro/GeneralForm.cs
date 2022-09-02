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
                + @"\Box\PM Resources\Archives\Layout Team 2022\Repos Macros\Blocks Files\";
        readonly int n = 26;
        private bool res;
        readonly int n2 = 26;
        private string num;
        private string type;
        bool active = true;
        readonly List<CheckBox> list_CheckBoxes = new List<CheckBox>();
        readonly List<CheckBox> list_CheckBoxesSize = new List<CheckBox>();
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
            try
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
                        string pathFile = pathGeneral + comboBox1.SelectedItem.ToString() + @"\" + type + @"\" + num + @"\";
                        for (int i = 0; i < checkedListBox1.Items.Count; i++)
                        {
                            if (checkedListBox1.CheckedItems.Contains(checkedListBox1.Items[i]))
                            {
                                var spl = checkedListBox1.Items[i].ToString().Split('\\');
                                string letter = spl.Last();
                                MyCommands myCommands = new MyCommands();
                                myCommands.CAST(pathFile, ppr, letter);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void GetFoldersType()
        {
            try
            {
                this.groupBox1.Controls.Clear();
                int m = 0;
                var subFolders = Directory.GetDirectories(pathGeneral + comboBox1.SelectedItem.ToString());
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void GetFoldersSize()
        {
            try
            {
                this.groupBox2.Controls.Clear();
                int m2 = 0;
                var subFolders = Directory.GetDirectories(pathGeneral + comboBox1.SelectedItem.ToString());
                var firstSubFolder = subFolders.First();
                string[] splited_0 = firstSubFolder.Split('\\');
                string sizePath = pathGeneral + comboBox1.SelectedItem.ToString() + "\\" + splited_0.Last().ToString();
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void GetFoldersTypeCleanUp()
        {
            this.groupBox1.Controls.Clear();
        }
        public void GetFoldersSizeCleanUp()
        {
            this.groupBox2.Controls.Clear();
        }
        private void Custom_event_handler(object sender, EventArgs e)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
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
            return "\\" + a + "\\" + b;
        }
        public void ChangeCheckBoxList()
        {
            try
            {
                var AB = VerifyCheckBoxList();
                var path = pathGeneral + comboBox1.SelectedItem.ToString();
                var fileNames = Directory.GetFiles(path + AB);
                checkedListBox1.Items.Clear();
                foreach (var fileName in fileNames)
                {
                    string[] splited = fileName.Split('\\');
                    checkedListBox1.Items.Add(splited.Last());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void Custom_event_handlerSize(object sender, EventArgs e)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void BtnBrowseFile_Click(object sender, EventArgs e)
        {
            try
            {
                using (var fbd = new FolderBrowserDialog())
                {
                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        pathGeneral = fbd.SelectedPath + @"\";
                        var subFolders = Directory.GetDirectories(fbd.SelectedPath);
                        comboBox1.Items.Clear();
                        foreach (var subFolder in subFolders)
                        {
                            string[] splited = subFolder.Split('\\');
                            comboBox1.Items.Add(splited.Last().ToString());
                        }
                        GetFoldersTypeCleanUp();
                        GetFoldersSizeCleanUp();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void GeneralForm_Activated(object sender, EventArgs e)
        {
            try
            {
                if (active)
                {
                    string pathFile2 = pathGeneral;
                    var subFolders = Directory.GetDirectories(pathFile2);
                    foreach (var subFolder in subFolders)
                    {
                        string[] splited2 = subFolder.Split('\\');
                        comboBox1.Items.Add(splited2.Last().ToString());
                    }
                    active = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void CheckedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            try
            {
                string itemText = checkedListBox1.Items[e.Index].ToString();
                var AB = VerifyCheckBoxList();
                var path = pathGeneral + comboBox1.SelectedItem.ToString();
                var totalFilePath = path + AB + "\\" + itemText;
                for (int ix = 0; ix < checkedListBox1.Items.Count; ++ix)
                {
                    if (ix != e.Index) checkedListBox1.SetItemChecked(ix, false);
                }
                if (e.NewValue == CheckState.Checked)
                {
                    _ = new ObjectIdCollection();
                    using (Database OuterDB = new Database())
                    {
                        OuterDB.ReadDwgFile(totalFilePath, System.IO.FileShare.Read, false, "");
                        using (Transaction tr = OuterDB.TransactionManager.StartTransaction())
                        {
                            BlockTable bt;
                            bt = (BlockTable)tr.GetObject(OuterDB.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord blk = (BlockTableRecord)tr.GetObject(bt["*Model_Space"], OpenMode.ForRead);
                            var imgsrc = Autodesk.AutoCAD.Windows.Data.CMLContentSearchPreviews.GetBlockTRThumbnail(blk);
                            var bmp = ImageSourceToGDI(imgsrc as System.Windows.Media.Imaging.BitmapSource);
                            pictureBox2.BackgroundImageLayout = ImageLayout.Stretch;
                            pictureBox2.BackgroundImage = bmp;
                            tr.Commit();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private static System.Drawing.Image ImageSourceToGDI(System.Windows.Media.Imaging.BitmapSource src)
        {
            var ms = new MemoryStream();
            var encoder = new System.Windows.Media.Imaging.BmpBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(src));
            encoder.Save(ms);
            ms.Flush();
            return System.Drawing.Image.FromStream(ms);
        }

        private void ComboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                string pathFile = pathGeneral + comboBox1.SelectedItem.ToString();
                string[] splited = pathFile.Split('\\');
                this.Text = splited.Last();
                GetFoldersType();
                GetFoldersSize();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
