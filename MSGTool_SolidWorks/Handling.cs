using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace MSGTool_SolidWorks {
    internal static class Handling {

        public static swDocumentTypes_e DocType { get; set; } = swDocumentTypes_e.swDocNONE;
        public static readonly ISldWorks _SldWorks = SWUtils.GetSwApp()?.Sw;
        public static bool CommandInProgress { get => _SldWorks.CommandInProgress; set => _SldWorks.CommandInProgress = value; }
        public static bool DocIsClose() {
            if(_SldWorks.ActiveDoc != null) {
                _SldWorks.SendMsgToUser("为了程序稳定运行请关闭所有打开文件！");
                return false;
            }
            return true;
        }
        public static List<string> FilePaths { get; set; } = new List<string>();
        public static List<string> Temps { get; set; } = new List<string>();
        public static string ConvertFiles(string filePath, swDocumentTypes_e swDocTypes, string extendName) {
            string Warning;

            try {
                // 打开工程图文件
                int fileError = 0;
                int fileWarning = 0;
                ModelDoc2 doc = _SldWorks.OpenDoc6(filePath, (int)swDocTypes, 1, "", ref fileError, ref fileWarning);
                if(doc == null) Warning = ($"打开工程图文件失败，错误码：{fileError}，警告码：{fileWarning}");
                else {
                    doc.SaveAs3(Path.ChangeExtension(filePath, extendName), 0, 1);
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

        public static string FileSwitch(string fileFullName, string extendName) {
            switch(DocType) {
                case swDocumentTypes_e.swDocNONE:
                    break;
                case swDocumentTypes_e.swDocPART:
                    return ConvertFiles(fileFullName, swDocumentTypes_e.swDocPART, extendName);
                case swDocumentTypes_e.swDocASSEMBLY:
                    break;
                case swDocumentTypes_e.swDocDRAWING:
                    return ConvertFiles(fileFullName, swDocumentTypes_e.swDocDRAWING, extendName);
                case swDocumentTypes_e.swDocSDM:
                    break;
                case swDocumentTypes_e.swDocLAYOUT:
                    break;
            }
            return string.Empty;
        }

        public static bool GetAsmDocs(HashSet<string> paths,out bool? existDrw) {

            existDrw = null;
            if(!(_SldWorks.ActiveDoc is AssemblyDoc asmSwDoc)) {
                _SldWorks.SendMsgToUser("当前没有活动文档或活动文档不是装配文档");
                return false;
            }

            var components = (object[])asmSwDoc.GetComponents(false);

            foreach(Component2 comp2 in components.Cast<Component2>()) {
                if(comp2 == null) continue;
                if(!(comp2.GetModelDoc2() is ModelDoc2 modelDoc)) continue;
                var partPth = modelDoc.GetPathName();
                if(!paths.Add(partPth)) continue;

                var docType = modelDoc.GetType();
                var drw = Path.ChangeExtension(partPth, "SLDDRW");

                switch(DocType) {
                    case swDocumentTypes_e.swDocDRAWING:
                        existDrw = null;
                        break;
                    case swDocumentTypes_e.swDocPART:
                        if(docType == 3) continue;
                        existDrw = File.Exists(drw);
                        break;
                    case swDocumentTypes_e.swDocASSEMBLY:
                        if(docType == 3) continue;
                        existDrw = File.Exists(drw);
                        break;
                    default:
                        continue;
                }
            }

            return paths.Count != 0;
        }

        public static void OpenDoc(string fileName) {

            // 获取所有打开的文档
            if(_SldWorks.GetDocuments() is object[] swObj) {
                List<string> openDocTitle = swObj.Select(o => (o as ModelDoc2).GetTitle()).ToList();
                if(openDocTitle.FindIndex(t => string.Equals(t, fileName, StringComparison.OrdinalIgnoreCase)) != -1) {
                    int _sld = 0;
                    _SldWorks.ActivateDoc3(Name: fileName, false, 1, ref _sld);
                    return;
                }
            }

            // 根据文件扩展名判断文件类型
            int docType;
            switch(Path.GetExtension(fileName).ToUpper()) {
                case SWUtils.PrtDoc:
                    docType = (int)swDocumentTypes_e.swDocPART;
                    break;
                case SWUtils.AsmDoc:
                    docType = (int)swDocumentTypes_e.swDocIMPORTED_ASSEMBLY;
                    break;
                case SWUtils.DrwDoc:
                    docType = (int)swDocumentTypes_e.swDocDRAWING;
                    break;
                default:
                    return;
            }

            int fileError = 0;
            int fileWarning = 0;
            _SldWorks.OpenDoc6(fileName, docType, 1, "", ref fileError, ref fileWarning);
        }

        public static void TraverseDirectory(string rootDir) {

            if(!Directory.Exists(rootDir)) return;

            Temps.Clear();
            Queue<DirectoryInfo> dirQueue = new Queue<DirectoryInfo>();
            dirQueue.Enqueue(new DirectoryInfo(rootDir));

            while(dirQueue.Count > 0) {
                DirectoryInfo currentDir = dirQueue.Dequeue();
                try {
                    foreach(FileInfo file in currentDir.GetFiles())
                        if(IsValidFile(file.FullName)) Temps.Add(file.FullName);

                    foreach(DirectoryInfo subDir in currentDir.GetDirectories())
                        dirQueue.Enqueue(subDir);

                } catch(UnauthorizedAccessException ex) {
                    System.Diagnostics.Debug.WriteLine($"无权限访问目录：{currentDir.FullName}，错误：{ex.Message}");
                } catch(PathTooLongException ex) {
                    System.Diagnostics.Debug.WriteLine($"路径过长：{currentDir.FullName}，错误：{ex.Message}");
                }
            }
        }

        private static bool IsValidFile(string filePath) {
            if(Path.GetFileName(filePath)[0] == '~') return false;
            var ext = Path.GetExtension(filePath).ToUpperInvariant();
            switch(DocType) {
                case swDocumentTypes_e.swDocDRAWING: return ext == SWUtils.DrwDoc;
                case swDocumentTypes_e.swDocPART: return ext == SWUtils.PrtDoc;
                case swDocumentTypes_e.swDocASSEMBLY: return ext == SWUtils.AsmDoc || ext == SWUtils.PrtDoc;
                default: return false;
            }
        }

        public static bool ExcludeDrwDoc(string file)
            => File.Exists(Path.ChangeExtension(file, "SLDDRW"));

    }
}
