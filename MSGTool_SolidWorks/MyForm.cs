using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace MSGTool_SolidWorks {

    public partial class MyForm : Form {

        public MyForm() => InitializeComponent();
        void MyForm_Load(object sender, EventArgs e) {

            if(Handling._SldWorks == null) {
                MessageBox.Show("请打开 solid works 程序！");
                this.Close();
                return;
            }

            SwDocType.Text = "工程图";
            ExtendName.Text = "Pdf";
            Handling.DocType = swDocumentTypes_e.swDocDRAWING;

            Bar.Visible = false;
            // 设置 ListView 的视图模式为详细信息视图，因为表头只有在详细信息视图中才会显示
            ListFile.View = System.Windows.Forms.View.Details;
            ListFile.GridLines = true;
            ListFile.HeaderStyle = ColumnHeaderStyle.Nonclickable; // 禁用表头点击排序
            ListFile.MultiSelect = false;
            ListFile.FullRowSelect = true;
            ListFile.AllowDrop = true;

            ListFile.OwnerDraw = false; // 开启自定义绘制
            // 添加表头列
            ListFile.Columns.Add("序号", 50);
            ListFile.Columns.Add("文件名", 120);
            ListFile.Columns.Add("同名工程图", 50);
            ListFile.Columns.Add("文件路径", 280);

            int totalWidthOfOtherColumns = 0;
            for(int i = 0; i < ListFile.Columns.Count - 1; i++) totalWidthOfOtherColumns += ListFile.Columns[i].Width;
            ListFile.Columns[ListFile.Columns.Count - 1].Width = ListFile.ClientSize.Width - totalWidthOfOtherColumns;
            // 处理窗体的 Resize 事件
            Resize += MyForm_Resize;
        }
        void MyForm_Resize(object sender, EventArgs e) {

            int lastColumnIndex = ListFile.Columns.Count - 1;
            if(lastColumnIndex < 0) return;

            int totalWidthOfOtherColumns = 0;
            for(int i = 0; i < lastColumnIndex; i++)
                totalWidthOfOtherColumns += ListFile.Columns[i].Width;

            ListFile.Columns[lastColumnIndex].Width = ListFile.ClientSize.Width - totalWidthOfOtherColumns;
        }
        private void Export_Click(object sender, EventArgs e) {

            if(!Handling.DocIsClose()) return;

            ListFile.BeginUpdate();
            ListFile.Columns.Add("处理结果", 280);

            Bar.Value = 0;
            Bar.Visible = true;
            int totalFiles = ListFile.Items.Count;
            int processedFiles = 0;

            foreach(ListViewItem item in ListFile.Items) {
                if(item.SubItems.Count < 2) {
                    MessageBox.Show("数据格式错误");
                    return;
                }
                processedFiles++;
                Bar.Value = (int)((double)processedFiles / totalFiles * 100);
                string fileFullName = item.SubItems[3].Text;
                var extendName = ExtendName.Text;

                item.SubItems.Add(Handling.FileSwitch(fileFullName, extendName));
            }

            ListFile.EndUpdate();
        }
        private void AddFolder_Click(object sender, EventArgs e) {

            using(var folderBrowserDialog = new FolderBrowserDialog { Description = "请选择一个文件夹" }) {
                DialogResult result = folderBrowserDialog.ShowDialog();
                if(result != DialogResult.OK) return;
                Handling.TraverseDirectory(folderBrowserDialog.SelectedPath);
            }

            ListFile.BeginUpdate();

            foreach(var file in Handling.Temps) {

                ListViewItem item1 = new ListViewItem((ListFile.Items.Count + 1).ToString());
                item1.SubItems.Add(Path.GetFileName(file));
                if(Handling.DocType == swDocumentTypes_e.swDocDRAWING) 
                    item1.SubItems.Add("");
                else 
                    item1.SubItems.Add(File.Exists(Path.ChangeExtension(file, "SLDDRW")).ToString());

                item1.SubItems.Add(file);

                ListFile.Items.Add(item1);
            }
           ListFile.EndUpdate();

        }
        private void Remove_Click(object sender, EventArgs e) {
            ListFile.BeginUpdate();
            foreach(ListViewItem item in ListFile.SelectedItems)
                ListFile.Items.Remove(item);
            ListFile.EndUpdate();
        }
        private void Clear_Click(object sender, EventArgs e) {
            ListFile.Items.Clear();
        }
        private void AddAsm_Click(object sender, EventArgs e) {

            var list = new List<ListViewItem>();
            Bar.Value = 0;
            Bar.Visible = true;

            //ListViewItem item1 = new ListViewItem(count.ToString());
            //item1.SubItems.Add(fileName);
            //item1.SubItems.Add(@bool.ToString());
            ////item1.SubItems.Add(docType.ToString());
            //item1.SubItems.Add(partPth);
            //list.Add(item1);

            //int processedFiles = 0;
            //processedFiles++;
            //progressBar1.Value = (int)((double)processedFiles / totalFiles * 100);

            ListFile.BeginUpdate();
            ListFile.Items.AddRange(list.ToArray());
            ListFile.EndUpdate();

            Bar.Visible = false;
        }
        private void SwOPen_Click(object sender, EventArgs e) {
            if(ListFile.Items.Count < 1) return;
            Handling.OpenDoc(ListFile.SelectedItems[0].SubItems[3].Text);
        }
        private void SwDocType_SelectedIndexChanged(object sender, EventArgs e) {
            string[] extends;
            switch(SwDocType.Text) {
                case "工程图":
                    extends = new string[] { "Pdf", "Dxf", "Dwg", "Png" };
                    Handling.DocType = swDocumentTypes_e.swDocDRAWING;

                    ListFile.BeginUpdate();

                    foreach(ListViewItem item in ListFile.Items) {
                        if(item.SubItems[2].Text == "") continue;
                        bool.TryParse(item.SubItems[2].Text, out var isDrw);
                        if(!isDrw) {
                            ListFile.Items.Remove(item);
                        } else {
                            item.SubItems[3].Text = Path.ChangeExtension(item.SubItems[3].Text, "SLDDRW");
                            item.SubItems[2].Text = "";
                            item.SubItems[1].Text = Path.ChangeExtension(item.SubItems[1].Text, "SLDDRW");
                        }
                    }

                    ListFile.EndUpdate();

                    break;
                case "零件":
                    extends = (new string[] { "Stp", "Step", "X_t", "Pdf", "Png" });
                    Handling.DocType = swDocumentTypes_e.swDocPART;
                    break;
                case "钣金":
                    extends = new string[] { "Dxf", "Stp", "Step", "X_t", "Pdf", "Png" };
                    Handling.DocType = swDocumentTypes_e.swDocPART;
                    break;
                case "装配|零件":
                    extends = new string[] { "Stp", "Step", "X_t", "Pdf", "Png" };
                    Handling.DocType = swDocumentTypes_e.swDocASSEMBLY;
                    break;
                default:
                    return;
            }
            ExtendName.Items.Clear();
            ExtendName.Items.AddRange(extends);
            ExtendName.Text = extends[0];
        }
        private void Options_Click(object sender, EventArgs e) {
            //设置选项增加属性过滤
        }
        private void ListFile_DragDrop(object sender, DragEventArgs e) {

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            Handling.Temps.Clear();
            foreach(string file in files)
                if(Directory.Exists(file))
                    Handling.TraverseDirectory(file);

            if(Handling.Temps.Count > 2000) {
                MessageBox.Show("文件数量超过2000");
                return;
            }

            ListFile.BeginUpdate();

            foreach(var file in Handling.Temps) {

                ListViewItem item = new ListViewItem((ListFile.Items.Count + 1).ToString());

                item.SubItems.Add(Path.GetFileName(file));

                if(Handling.DocType == swDocumentTypes_e.swDocDRAWING)
                    item.SubItems.Add("");
                else
                    item.SubItems.Add(Handling.ExcludeDrwDoc(file).ToString());

                item.SubItems.Add(file);

                ListFile.Items.Add(item);
            }

            ListFile.EndUpdate();
        }
        private void ListFile_DragEnter(object sender, DragEventArgs e) {
            if(e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

    }

}
