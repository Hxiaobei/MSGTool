using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using System.Threading;

namespace MSGTool_SolidWorks {

    public partial class Form_ExportTool : Form {

        static ISldWorks _SldWorks;
        static SwDocumentTypes documentTypes = SwDocumentTypes.swDocNONE;
        public Form_ExportTool() => InitializeComponent();

        void Form_ExportTool_Load(object sender, EventArgs e) {

            _SldWorks = SWUtils.GetSwApp()?.Sw;
            if(_SldWorks == null) {
                MessageBox.Show("请打开 solid works 程序！");
                this.Close();
                return;
            }

            toolStripComboBox1.Text = "工程图";
            toolStripComboBox2.Text = "Pdf";
            documentTypes = SwDocumentTypes.swDocDRAWING;

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
            if(_SldWorks.ActiveDoc != null) _SldWorks.SendMsgToUser("为了程序稳定运行请关闭所有打开文件！");
            _SldWorks.CommandInProgress = true;

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
                string wr;
                switch(documentTypes) {
                    case SwDocumentTypes.swDocNONE:
                        break;
                    case SwDocumentTypes.swDocPART:
                        wr = ConvertFlie(fileFullName, swDocumentTypes_e.swDocPART, toolStripComboBox2.Text);
                        item.SubItems.Add(wr);
                        break;
                    case SwDocumentTypes.swDocASSEMBLY:
                        break;
                    case SwDocumentTypes.swDocDRAWING:
                        wr = ConvertFlie(fileFullName, swDocumentTypes_e.swDocDRAWING, toolStripComboBox2.Text);
                        item.SubItems.Add(wr);
                        break;
                    case SwDocumentTypes.swDocSDM:
                        break;
                    case SwDocumentTypes.swDocLAYOUT:
                        break;
                    case SwDocumentTypes.swDocSM:
                        break;
                    default:
                        break;
                }
            }


            listView1.EndUpdate();

            _SldWorks.CommandInProgress = false;
        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e) {
            // 创建 FolderBrowserDialog 对象
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            // 设置对话框的属性（可选）
            folderBrowserDialog.Description = "请选择一个文件夹";

            // 显示对话框
            DialogResult result = folderBrowserDialog.ShowDialog();

            // 处理用户的选择
            if(result != DialogResult.OK) return;

            string selectedFolder = folderBrowserDialog.SelectedPath;
            List<string> list = new List<string>();

            TraverseDirectory(selectedFolder, list);
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
                if(documentTypes == SwDocumentTypes.swDocDRAWING) {
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

        }
        //清空列表
        private void toolStripMenuItem3_Click(object sender, EventArgs e) {
            // 清空 ListView 中的所有数据项
            listView1.Items.Clear();
        }

        private void 添加当前装配ToolStripMenuItem_Click(object sender, EventArgs e) {

            if(!(_SldWorks.ActiveDoc is AssemblyDoc asmSwDoc)) {
                _SldWorks.SendMsgToUser("当前没有活动文档或活动文档不是装配文档");
                return;
            }

            _SldWorks.CommandInProgress = true;
            //顶层 true 所有 false
            var components = (object[])asmSwDoc.GetComponents(false);
            var count = 1;
            var list = new List<ListViewItem>();
            bool? @bool;
            bool isAdd;

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            int totalFiles = components.Length;
            int processedFiles = 0;
            foreach(Component2 comp2 in components.Cast<Component2>()) {
                processedFiles++;
                progressBar1.Value = (int)((double)processedFiles / totalFiles * 100);

                isAdd = true;
                if(comp2 == null) continue;
                if(!(comp2.GetModelDoc2() is ModelDoc2 modelDoc)) continue;
                var partPth = modelDoc.GetPathName();
                var fileName = Path.GetFileName(partPth);
                foreach(ListViewItem item in list) {
                    if(string.Equals(item.SubItems[1].Text, fileName, StringComparison.OrdinalIgnoreCase)) {
                        isAdd = false;
                        break;
                    }
                }
                if(!isAdd) continue;
                var docType = modelDoc.GetType();
                var drw = Path.ChangeExtension(partPth, "SLDDRW");
                switch(documentTypes) {
                    case SwDocumentTypes.swDocDRAWING:
                        if(!File.Exists(drw)) continue;
                        @bool = null;
                        break;
                    case SwDocumentTypes.swDocPART:
                        if(docType == 2) continue;
                        @bool = File.Exists(drw);
                        break;
                    case SwDocumentTypes.swDocSM:
                        if(docType == 2) continue;
                        @bool = File.Exists(drw);
                        break;
                    default:
                        @bool = File.Exists(drw);
                        break;
                }

                ListViewItem item1 = new ListViewItem(count.ToString());
                item1.SubItems.Add(fileName);
                item1.SubItems.Add(@bool.ToString());
                //item1.SubItems.Add(docType.ToString());
                item1.SubItems.Add(partPth);
                list.Add(item1);

                count++;
            }


            _SldWorks.CommandInProgress = false;

            listView1.BeginUpdate();
            listView1.Items.AddRange(list.ToArray());
            listView1.EndUpdate();

            progressBar1.Visible = false;
        }

        // solid works 中打开
        private void solidworksToolStripMenuItem_Click(object sender, EventArgs e) {
            if(listView1.Items.Count < 1) return;
            string subItemText = listView1.SelectedItems[0].SubItems[1].Text;
            // 获取所有打开的文档
            if(_SldWorks.GetDocuments() is object[] swObj) {
                List<string> openDocTitle = swObj.Select(o => (o as ModelDoc2).GetTitle()).ToList();
                if(openDocTitle.FindIndex(t => string.Equals(t, subItemText, StringComparison.OrdinalIgnoreCase)) != -1) {
                    int _sld = 0;
                    _SldWorks.ActivateDoc3(Name: subItemText, false, 1, ref _sld);
                    return;
                }
            }

            var filePath = listView1.SelectedItems[0].SubItems[3].Text;

            // 根据文件扩展名判断文件类型
            int docType;
            switch(Path.GetExtension(filePath).ToLowerInvariant()) {
                case ".sldprt":
                    docType = (int)swDocumentTypes_e.swDocPART;
                    break;
                case ".sldasm":
                    docType = (int)swDocumentTypes_e.swDocIMPORTED_ASSEMBLY;
                    break;
                case ".slddrw":
                    docType = (int)swDocumentTypes_e.swDocDRAWING;
                    break;

                default:
                    return;
            }

            int fileError = 0;
            int fileWarning = 0;
            _SldWorks.OpenDoc6(filePath, docType, 1, "", ref fileError, ref fileWarning);

        }

        //修改导出文档类型//(*.SLDASM),(*.SLDPRT),(*.SLDDRW)
        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            switch(toolStripComboBox1.Text) {
                case "工程图":
                    toolStripComboBox2.Items.Clear();
                    toolStripComboBox2.Items.AddRange(new object[] { "Pdf", "Dxf", "Dwg", "Png" });
                    toolStripComboBox2.Text = "Pdf";
                    documentTypes = SwDocumentTypes.swDocDRAWING;
                    listView1.BeginUpdate();
                    foreach(ListViewItem item in listView1.Items) {
                        var isDrw = item.SubItems[2].Text;
                        if(isDrw == false.ToString()) {
                            listView1.Items.Remove(item);
                        } else if(isDrw == true.ToString()) {
                            var filePath = item.SubItems[3].Text;
                            var fileName = item.SubItems[1].Text;
                            item.SubItems[3].Text = Path.ChangeExtension(filePath, "SLDDRW");
                            item.SubItems[2].Text = "";
                            item.SubItems[1].Text = Path.ChangeExtension(fileName, "SLDDRW");
                        }

                    }
                    listView1.EndUpdate();
                    break;
                case "零件":
                    toolStripComboBox2.Items.Clear();
                    toolStripComboBox2.Items.AddRange(new object[] { "Stp", "Step", "X_t", "Pdf" });
                    toolStripComboBox2.Text = "Step";
                    documentTypes = SwDocumentTypes.swDocPART;
                    break;
                case "钣金":
                    toolStripComboBox2.Items.Clear();
                    toolStripComboBox2.Items.AddRange(new object[] { "Dxf", "Stp", "Step", "X_t", "Pdf" });
                    toolStripComboBox2.Text = "Dxf";
                    documentTypes = SwDocumentTypes.swDocSM;
                    break;
                case "装配|零件":
                    toolStripComboBox2.Items.Clear();
                    toolStripComboBox2.Items.AddRange(new object[] { "Stp", "Step", "X_t", "Pdf" });
                    toolStripComboBox2.Text = "Step";
                    documentTypes = SwDocumentTypes.swDocASSEMBLY;
                    break;
            }
        }

        /// <summary>
        /// 设置选项增加属性过滤
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 设置ToolStripMenuItem_Click(object sender, EventArgs e) {

        }

        static void TraverseDirectory(string dir, List<string> list) {
            DirectoryInfo directoryInfo = new DirectoryInfo(dir);
            FileInfo[] files = directoryInfo.GetFiles();
            DirectoryInfo[] subDirectories = directoryInfo.GetDirectories();

            foreach(FileInfo file in files) {
                var fileFullName = file.FullName;
                if(fileFullName.IndexOf('~') != -1) continue;
                var ext = Path.GetExtension(fileFullName).ToUpper();
                switch(documentTypes) {
                    case SwDocumentTypes.swDocDRAWING:
                        if(ext == ".SLDDRW") list.Add(fileFullName);
                        break;
                    case SwDocumentTypes.swDocPART:
                    case SwDocumentTypes.swDocSM:
                        if(ext == ".SLDPRT") list.Add(fileFullName);
                        break;
                    default:
                        if(ext == ".SLDASM") list.Add(fileFullName);
                        else if(ext == ".SLDPRT") list.Add(fileFullName);
                        break;
                }
            }

            foreach(DirectoryInfo subDirectory in subDirectories) {
                TraverseDirectory(subDirectory.FullName, list);
            }
        }

        private void listView1_DragDrop(object sender, DragEventArgs e) {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            List<string> list = new List<string>();

            foreach(string file in files) {
                if(Directory.Exists(file)) {
                    TraverseDirectory(file, list);
                } else if(File.Exists(file)) {
                    if(file.IndexOf('~') != -1) continue;
                    var ext = Path.GetExtension(file).ToUpper();
                    switch(documentTypes) {
                        case SwDocumentTypes.swDocDRAWING:
                            if(ext == ".SLDDRW") list.Add(file);
                            break;
                        case SwDocumentTypes.swDocPART:
                        case SwDocumentTypes.swDocSM:
                            if(ext == ".SLDPRT") list.Add(file);
                            break;
                        default:
                            if(ext == "SLDASM") list.Add(file);
                            else if(ext == "SLDPRT") list.Add(file);
                            break;
                    }
                }
            }

            if(list.Count > 2000) {
                MessageBox.Show("文件数量超过2000");
                return;
            }

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
                if(documentTypes == SwDocumentTypes.swDocDRAWING) {
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

        private void listView1_DragEnter(object sender, DragEventArgs e) {
            if(e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        #region Code solid works
        static string ConvertFlie(string filePath, swDocumentTypes_e swDocTypes, string extenName) {
            string Warning;

            try {
                // 打开工程图文件
                int fileError = 0;
                int fileWarning = 0;
                ModelDoc2 doc = _SldWorks.OpenDoc6(filePath, (int)swDocTypes, 1, "", ref fileError, ref fileWarning);
                if(doc == null) Warning = ($"打开工程图文件失败，错误码：{fileError}，警告码：{fileWarning}");
                else {
                    doc.SaveAs3(Path.ChangeExtension(filePath, extenName), 0, 1);
                    Warning = "成功";
                    // 关闭文档
                    _SldWorks.CloseDoc(doc.GetTitle());
                }
            } catch(COMException ex) {
                Warning = ($"COM 异常：{ex.Message}");
            } catch(Exception ex) {
                Warning = ($"其他异常：{ex.Message}");
            }

            return Warning;
        }

        #endregion
    }
}
