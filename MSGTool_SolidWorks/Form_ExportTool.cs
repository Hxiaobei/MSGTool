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

    public partial class Form_ExportTool : Form {

        public Form_ExportTool() => InitializeComponent();

        void Form_ExportTool_Load(object sender, EventArgs e) {

            if(Handling._SldWorks == null) {
                MessageBox.Show("请打开 solid works 程序！");
                this.Close();
                return;
            }

            toolStripComboBox1.Text = "工程图";
            toolStripComboBox2.Text = "Pdf";
            Handling.DocType = swDocumentTypes_e.swDocDRAWING;

            progressBar1.Visible = false;
            // 设置 ListView 的视图模式为详细信息视图，因为表头只有在详细信息视图中才会显示
            listView1.View = System.Windows.Forms.View.Details;
            listView1.GridLines = true;
            listView1.HeaderStyle = ColumnHeaderStyle.Nonclickable; // 禁用表头点击排序
            listView1.MultiSelect = false;
            listView1.FullRowSelect = true;
            listView1.AllowDrop = true;

            listView1.OwnerDraw = false; // 开启自定义绘制
            // 添加表头列
            listView1.Columns.Add("序号", 50);
            listView1.Columns.Add("文件名", 120);
            listView1.Columns.Add("同名工程图", 50);
            listView1.Columns.Add("文件路径", 280);

            int totalWidthOfOtherColumns = 0;
            for(int i = 0; i < listView1.Columns.Count - 1; i++) totalWidthOfOtherColumns += listView1.Columns[i].Width;
            listView1.Columns[listView1.Columns.Count - 1].Width = listView1.ClientSize.Width - totalWidthOfOtherColumns;
            // 处理窗体的 Resize 事件
            Resize += Form1_Resize;
        }

        void Form1_Resize(object sender, EventArgs e) {
            // 确保 ListView 至少有一列
            if(listView1.Columns.Count < 1) return;

            // 获取最后一列的索引
            int lastColumnIndex = listView1.Columns.Count - 1;

            // 计算所有列（除最后一列）的总宽度
            int totalWidthOfOtherColumns = 0;
            for(int i = 0; i < listView1.Columns.Count - 1; i++)
                totalWidthOfOtherColumns += listView1.Columns[i].Width;


            // 计算 ListView 的可用宽度
            int availableWidth = listView1.ClientSize.Width;

            // 设置最后一列的宽度为剩余宽度
            listView1.Columns[lastColumnIndex].Width = availableWidth - totalWidthOfOtherColumns;
        }

        private void 导出ToolStripMenuItem_Click(object sender, EventArgs e) {

            if(!Handling.DocIsClose()) return;

            listView1.BeginUpdate();
            listView1.Columns.Add("处理结果", 280);

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            int totalFiles = listView1.Items.Count;
            int processedFiles = 0;

            foreach(ListViewItem item in listView1.Items) {
                if(item.SubItems.Count < 2) {
                    MessageBox.Show("数据格式错误");
                    return;
                }
                processedFiles++;
                progressBar1.Value = (int)((double)processedFiles / totalFiles * 100);
                string fileFullName = item.SubItems[3].Text;
                var extendName = toolStripComboBox2.Text;

                item.SubItems.Add(Handling.FileSwitch(fileFullName, extendName));
            }

            listView1.EndUpdate();
        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e) {

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog { Description = "请选择一个文件夹" };
            DialogResult result = folderBrowserDialog.ShowDialog();

            // 处理用户的选择
            if(result != DialogResult.OK) return;

            string selectedFolder = folderBrowserDialog.SelectedPath;
            List<string> list = new List<string>();

            Handling.TraverseDirectory(selectedFolder);
            listView1.BeginUpdate();
            bool? @bool = null;
            var count = listView1.Items.Count + 1;
            bool isAdd;
            foreach(var file in list) {
                isAdd = true;
                foreach(ListViewItem item in listView1.Items) {
                    if(string.Equals(item.SubItems[3].Text, file, StringComparison.OrdinalIgnoreCase)) {
                        isAdd = false;
                        break;
                    }
                }
                if(!isAdd) continue;
                ListViewItem item1 = new ListViewItem(count.ToString());
                item1.SubItems.Add(Path.GetFileName(file));
                if(Handling.DocType == swDocumentTypes_e.swDocDRAWING) {
                    item1.SubItems.Add(@bool.ToString());
                } else {
                    var partPth = Path.ChangeExtension(file, "SLDDRW");
                    @bool = File.Exists(partPth);
                    item1.SubItems.Add(@bool.ToString());
                }
                item1.SubItems.Add(file);
                listView1.Items.Add(item1);
                count++;
            }
            listView1.EndUpdate();

        }

        //移除选择项
        private void toolStripMenuItem2_Click(object sender, EventArgs e) {
            listView1.BeginUpdate();
            foreach(ListViewItem item in listView1.SelectedItems)
                listView1.Items.Remove(item);
            listView1.EndUpdate();
        }
        //清空列表
        private void toolStripMenuItem3_Click(object sender, EventArgs e) {
            listView1.Items.Clear();
        }

        private void 添加当前装配ToolStripMenuItem_Click(object sender, EventArgs e) {

            var list = new List<ListViewItem>();
            progressBar1.Value = 0;
            progressBar1.Visible = true;

            //ListViewItem item1 = new ListViewItem(count.ToString());
            //item1.SubItems.Add(fileName);
            //item1.SubItems.Add(@bool.ToString());
            ////item1.SubItems.Add(docType.ToString());
            //item1.SubItems.Add(partPth);
            //list.Add(item1);

            //int processedFiles = 0;
            //processedFiles++;
            //progressBar1.Value = (int)((double)processedFiles / totalFiles * 100);

            listView1.BeginUpdate();
            listView1.Items.AddRange(list.ToArray());
            listView1.EndUpdate();

            progressBar1.Visible = false;
        }

        private void solidworksToolStripMenuItem_Click(object sender, EventArgs e) {
            if(listView1.Items.Count < 1) return;
            Handling.OpenDoc(listView1.SelectedItems[0].SubItems[3].Text);
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e) {

            string[] extends;
            switch(toolStripComboBox1.Text) {
                case "工程图":
                    extends = new string[] { "Pdf", "Dxf", "Dwg", "Png" };
                    Handling.DocType = swDocumentTypes_e.swDocDRAWING;

                    listView1.BeginUpdate();

                    foreach(ListViewItem item in listView1.Items) {
                        if(item.SubItems[2].Text == "") continue;
                        bool.TryParse(item.SubItems[2].Text, out var isDrw);
                        if(!isDrw) {
                            listView1.Items.Remove(item);
                        } else {
                            item.SubItems[3].Text = Path.ChangeExtension(item.SubItems[3].Text, "SLDDRW");
                            item.SubItems[2].Text = "";
                            item.SubItems[1].Text = Path.ChangeExtension(item.SubItems[1].Text, "SLDDRW");
                        }
                    }

                    listView1.EndUpdate();

                    break;
                case "零件":
                    extends = (new string[] { "Stp", "Step", "X_t", "Pdf" });
                    Handling.DocType = swDocumentTypes_e.swDocPART;
                    break;
                case "钣金":
                    extends = new string[] { "Dxf", "Stp", "Step", "X_t", "Pdf" };
                    Handling.DocType = swDocumentTypes_e.swDocPART;
                    break;
                case "装配|零件":
                    extends = new string[] { "Stp", "Step", "X_t", "Pdf" };
                    Handling.DocType = swDocumentTypes_e.swDocASSEMBLY;
                    break;
                default:
                    return;
            }
            toolStripComboBox2.Items.Clear();
            toolStripComboBox2.Items.AddRange(extends);
            toolStripComboBox2.Text = extends[0];
        }

        private void 设置ToolStripMenuItem_Click(object sender, EventArgs e) {
            //设置选项增加属性过滤
        }

        //接收拖住拽文件
        private void listView1_DragDrop(object sender, DragEventArgs e) {

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            Handling.Temps.Clear();
            foreach(string file in files)
                if(Directory.Exists(file))
                    Handling.TraverseDirectory(file);

            if(Handling.Temps.Count > 2000) {
                MessageBox.Show("文件数量超过2000");
                return;
            }

            listView1.BeginUpdate();

            foreach(var file in Handling.Temps) {

                ListViewItem item = new ListViewItem((listView1.Items.Count + 1).ToString());

                item.SubItems.Add(Path.GetFileName(file));

                if(Handling.DocType == swDocumentTypes_e.swDocDRAWING)
                    item.SubItems.Add("");
                else
                    item.SubItems.Add(Handling.ExcludeDrwDoc(file).ToString());

                item.SubItems.Add(file);

                listView1.Items.Add(item);
            }

            listView1.EndUpdate();
        }

        //接收拖住拽文件
        private void listView1_DragEnter(object sender, DragEventArgs e) {
            if(e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

    }

}
