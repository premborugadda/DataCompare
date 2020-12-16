using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataCompare.Functions
{
    class ReadFileNames
    {

        public void getSourceFileNames()
        {

            var directoryDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "Select Folder"
            };
            directoryDialog.ShowDialog();
            var homeDir = directoryDialog.SelectedFileTypeIndex;

        }
        // ...

        //OpenFileDialog folderBrowser = new OpenFileDialog();

        //folderBrowser.ValidateNames = false;
        //    folderBrowser.CheckFileExists = false;
        //    folderBrowser.CheckPathExists = true;
        //    // Always default to Folder Selection.
        //    folderBrowser.FileName = "Folder Selection.";
        //    if (folderBrowser.ShowDialog() == DialogResult.OK)
        //    {
        //        string folderPath = Path.GetDirectoryName(folderBrowser.FileName);

        //}



    //var fbd = new FolderBrowserDialog();
    //{
    //    DialogResult result = fbd.ShowDialog();

    //    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
    //    {
    //        string[] files = Directory.GetFiles(fbd.SelectedPath);

    //    System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
    //    }
    }
}

