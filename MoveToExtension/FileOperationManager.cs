using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MoveToExtension
{
    public static class FileOperationManager
    {
        [DllImport("shell32.dll")]
        private static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);
        private struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            public WFUNC wFunc;
            public string pFrom;
            public string pTo;
            public FILEOP_FLAGS fFlags;
            public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            public string lpszProgressTitle;
        }
        private enum WFUNC
        {
            FO_MOVE = 0x0001,
            FO_COPY = 0x0002,
            FO_DELETE = 0x0003,
            FO_RENAME = 0x0004
        }
        private enum FILEOP_FLAGS
        {
            FOF_MULTIDESTFILES = 0x0001, //pTo 指定了多个目标文件，而不是单个目录
            FOF_CONFIRMMOUSE = 0x0002,
            FOF_SILENT = 0x0044, // 不显示一个进度对话框 
            FOF_RENAMEONCOLLISION = 0x0008, // 碰到有抵触的名字时，自动分配前缀
            FOF_NOCONFIRMATION = 0x10, // 不对用户显示提示
            FOF_WANTMAPPINGHANDLE = 0x0020, // 填充 hNameMappings 字段，必须使用 SHFreeNameMappings 释放
            FOF_ALLOWUNDO = 0x40, // 允许撤销
            FOF_FILESONLY = 0x0080, // 使用 *.* 时, 只对文件操作
            FOF_SIMPLEPROGRESS = 0x0100, // 简单进度条，意味者不显示文件名。
            FOF_NOCONFIRMMKDIR = 0x0200, // 建新目录时不需要用户确定
            FOF_NOERRORUI = 0x0400, // 不显示出错用户界面
            FOF_NOCOPYSECURITYATTRIBS = 0x0800, // 不复制 NT 文件的安全属性
            FOF_NORECURSION = 0x1000 // 不递归目录
        }

        public static bool Copy(string path, string destdir, out string msg)
        => Copy(new string[] { path }, destdir, out msg);

        public static bool Copy(string[] paths, string destdir, out string msg)
        {
            return Operation(WFUNC.FO_COPY, paths, destdir, out msg);
        }

        public static bool MoveTo(string path, string destdir, out string msg)
         => MoveTo(new string[] { path }, destdir, out msg);

        public static bool MoveTo(string[] paths, string destdir, out string msg)
        {
            return Operation(WFUNC.FO_MOVE, paths, destdir, out msg);
        }

        public static bool Delete(string path, out string msg)
         => Delete(new string[] { path }, out msg);

        public static bool Delete(string[] paths, out string msg)
        {
            return Operation(WFUNC.FO_DELETE, paths, null, out msg);
        }

        public static bool Rename(string path, string destdir, out string msg)
         => Delete(new string[] { path }, out msg);

        public static bool Rename(string[] paths, string destdir, out string msg)
        {
            return Operation(WFUNC.FO_RENAME, paths, destdir, out msg);
        }

        private static bool Operation(WFUNC func, string[] paths, string targetDir, out string msg)
        {
            try
            {
                if (paths == null || paths.Length <= 0) throw new Exception("请选择需要操作的文件");
                if (string.IsNullOrWhiteSpace(targetDir))  throw new Exception("请选择目标目录");
                int i = 0, state=0;
                string pfrom = "", pto = "";
                msg = "";

                while (i < paths.Length)
                {
                    pfrom += paths[i] + "\x00";
                    pto += Path.Combine(targetDir, Path.GetFileName(paths[i])) + "\x00";

                    if (i % 100 == 0)
                    {
                        pfrom = pfrom + "\0";
                        pto = pto + "\0";
                        var lpFileOp = new SHFILEOPSTRUCT
                        {
                            wFunc = func,
                            pFrom = pfrom,
                            pTo = pto,
                            fFlags = FILEOP_FLAGS.FOF_MULTIDESTFILES | FILEOP_FLAGS.FOF_ALLOWUNDO| FILEOP_FLAGS.FOF_NOCONFIRMMKDIR,
                            fAnyOperationsAborted = false,
                        };
                        state = SHFileOperation(ref lpFileOp);
                        pfrom = "";
                        pto = "";
                        if(state != 0)
                        {
                            msg = GetAbortedString(state);
                            break;
                        }
                    }
                    i++;
                }
                if(state == 0 && !string.IsNullOrWhiteSpace(pfrom))
                {
                    pfrom = pfrom + "\0";
                    pto = pto + "\0";
                    var lpFileOp = new SHFILEOPSTRUCT
                    {
                        wFunc = func,
                        pFrom = pfrom,
                        pTo = pto,
                        fFlags = FILEOP_FLAGS.FOF_MULTIDESTFILES | FILEOP_FLAGS.FOF_ALLOWUNDO | FILEOP_FLAGS.FOF_NOCONFIRMMKDIR,
                        fAnyOperationsAborted = false,
                    };
                    state = SHFileOperation(ref lpFileOp);
                    msg = GetAbortedString(state);
                }

                return state == 0;
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                return false;
            }
        }

        private static string GetAbortedString(int code)
        {
            if (code == 0)
            {
                return "【操作完成】";
            }
            string state = code.ToString("X").ToUpper();
            switch (state)
            {
                case "4C7":
                    return "【您已取消操作】";
                case "74":
                    return "The source is a root directory, which cannot be moved or renamed.";
                case "76":
                    return "Security settings denied access to the source.";
                case "7C":
                    return "【路径错误】您操作路径的源或目标对象没有找到";
                case "10000":
                    return "An unspecified error occurred on the destination.";
                case "402":
                    return "An unknown error occurred. This is typically due to an invalid path " + "in the source or destination. This error does not occur on Windows Vista and later.";
                default:
                    return state;
            }
        }
    }

    public class FileOperationWrapper : IDisposable
    {
        private IFileOperation fileOperation;
        [ComImport]
        [Guid("947aab5f-0a5c-4c13-b4d6-4bf7836fc9f8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOperation
        {
            int Advise(IFileOperationProgressSink pfops, out int pdwCookie);
            int Unadvise(int dwCookie);
            int SetOperationFlags(uint dwOperationFlags);
            int SetProgressMessage([MarshalAs(UnmanagedType.LPWStr)] string pszMessage);
            int SetProgressDialog([MarshalAs(UnmanagedType.Interface)] object popd);
            int SetProperties([MarshalAs(UnmanagedType.Interface)] object pproparray);
            int SetOwnerWindow(uint hwndParent);
            int ApplyPropertiesToItem([MarshalAs(UnmanagedType.Interface)] object psiItem);
            int ApplyPropertiesToItems([MarshalAs(UnmanagedType.Interface)] object punkItems);
            int RenameItem([MarshalAs(UnmanagedType.Interface)] object psiItem,
                          [MarshalAs(UnmanagedType.LPWStr)] string pszNewName,
                          IFileOperationProgressSink pfopsItem);
            int RenameItems([MarshalAs(UnmanagedType.Interface)] object pUnkItems,
                           [MarshalAs(UnmanagedType.LPWStr)] string pszNewName);
            int MoveItem([MarshalAs(UnmanagedType.Interface)] object psiItem,
                        [MarshalAs(UnmanagedType.Interface)] object psiDestinationFolder,
                        [MarshalAs(UnmanagedType.LPWStr)] string pszNewName,
                        IFileOperationProgressSink pfopsItem);
            int MoveItems([MarshalAs(UnmanagedType.Interface)] object punkItems,
                         [MarshalAs(UnmanagedType.Interface)] object psiDestinationFolder);
            int CopyItem([MarshalAs(UnmanagedType.Interface)] object psiItem,
                        [MarshalAs(UnmanagedType.Interface)] object psiDestinationFolder,
                        [MarshalAs(UnmanagedType.LPWStr)] string pszCopyName,
                        IFileOperationProgressSink pfopsItem);
            int CopyItems([MarshalAs(UnmanagedType.Interface)] object punkItems,
                         [MarshalAs(UnmanagedType.Interface)] object psiDestinationFolder);
            int DeleteItem([MarshalAs(UnmanagedType.Interface)] object psiItem,
                          IFileOperationProgressSink pfopsItem);
            int DeleteItems([MarshalAs(UnmanagedType.Interface)] object punkItems);
            int NewItem([MarshalAs(UnmanagedType.Interface)] object psiDestinationFolder,
                       uint dwFileAttributes,
                       [MarshalAs(UnmanagedType.LPWStr)] string pszName,
                       [MarshalAs(UnmanagedType.LPWStr)] string pszTemplateName,
                       IFileOperationProgressSink pfopsItem);
            int PerformOperations();
            [PreserveSig]
            int GetAnyOperationsAborted(out bool pfAnyOperationsAborted);
        }

        [ComImport]
        [Guid("04B0F1A7-9490-44BC-96E1-4296A31252E2")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IFileOperationProgressSink
        {
        }

        public FileOperationWrapper()
        {
            Type type = Type.GetTypeFromCLSID(new Guid("3ad05575-8857-4850-9277-11b85bdb8e09"));
            fileOperation = (IFileOperation)Activator.CreateInstance(type);
        }

        public void CopyFile(string sourcePath1, string destPath)
        {
            IShellItem sourceItem1 = CreateShellItem(sourcePath1);
            IShellItem destFolderItem = CreateShellItem(System.IO.Path.GetDirectoryName(destPath));

            fileOperation.CopyItem(sourceItem1, destFolderItem,
                                 System.IO.Path.GetFileName(destPath), null);
            fileOperation.PerformOperations();

            Marshal.ReleaseComObject(sourceItem1);
            Marshal.ReleaseComObject(destFolderItem);
        }

        private IShellItem CreateShellItem(string path)
        {
            IntPtr psi;
            SHCreateItemFromParsingName(path, IntPtr.Zero,
                                      new Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"),
                                      out psi);
            return (IShellItem)Marshal.GetObjectForIUnknown(psi);
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            out IntPtr ppv);

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid,
                             [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(uint sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        public void Dispose()
        {
            if (fileOperation != null)
            {
                Marshal.ReleaseComObject(fileOperation);
                fileOperation = null;
            }
        }
    }
}
