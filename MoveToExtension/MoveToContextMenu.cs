using SharpShell;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[ComVisible(true)]
[COMServerAssociation(AssociationType.AllFiles)] // 关联所有文件类型
public class MoveToContextMenu : SharpContextMenu
{
    protected override bool CanShowMenu()
    {
        return true; // 对所有文件显示菜单
    }

    protected override ContextMenuStrip CreateMenu()
    {
        var menu = new ContextMenuStrip();

        // 获取当前选中文件的路径
        var selectedFilePath = SelectedItemPaths.First();

        // 创建"移动到"主菜单项
        var moveToMenuItem = new ToolStripMenuItem
        {
            Text = "移动到"
        };
        int i = 1;
        while (true)
        {
            try
            {
                var targetDir = GetParentDirectory(selectedFilePath, i + 1);
                if (targetDir == null) break;

                var dirName = new DirectoryInfo(targetDir).Name;
                var menuItem = new ToolStripMenuItem
                {
                    Text = $"{dirName} (上{i}层)",
                    Tag = targetDir // 存储目标路径
                };

                menuItem.Click += (sender, args) =>
                {
                    foreach (var p in SelectedItemPaths)
                    {
                        MoveFile(p, targetDir);
                    }
                };

                moveToMenuItem.DropDownItems.Add(menuItem);
                i++;
            }
            catch { break; }
        }

        // 添加根目录选项
        var rootPath = Path.GetPathRoot(selectedFilePath);
        var rootMenuItem = new ToolStripMenuItem
        {
            Text = $"{new DirectoryInfo(rootPath).Name} (根目录)",
            Tag = rootPath
        };
        rootMenuItem.Click += (sender, args) =>
        {
            foreach (var p in SelectedItemPaths)
            {
                MoveFile(p, rootPath);
            }
        };

        moveToMenuItem.DropDownItems.Add(new ToolStripSeparator());
        moveToMenuItem.DropDownItems.Add(rootMenuItem);

        menu.Items.Add(moveToMenuItem);
        return menu;
    }

    private string GetParentDirectory(string path, int levels)
    {
        string parent = path;
        for (int i = 0; i < levels; i++)
        {
            parent = Directory.GetParent(parent)?.FullName;
            if (parent == null) return null;
        }
        return parent;
    }

    private void MoveFile(string sourcePath, string targetDir)
    {
        try
        {
            var fileName = Path.GetFileName(sourcePath);
            var targetPath = Path.Combine(targetDir, fileName);

            File.Move(sourcePath, targetPath);
            Console.WriteLine($"文件已移动到: {targetPath}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"移动文件失败: {ex.Message}");
        }
    }
}

