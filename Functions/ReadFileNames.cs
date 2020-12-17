using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataCompare.Functions
{
    class ReadFileNames
    {

        public string[] getSourceFileNames(string homedir)
        {

            string[] fileEntries = Directory.GetFiles(homedir);
            string[] fileNames = {"","","","",""};
            foreach (string fileName in fileEntries)
            {
                if (fileName.Contains("~$") == false)
                {
                    if (fileName.ToUpper().Contains("KPI_EXTRACT_FULL"))
                    {
                        fileNames[0] = fileName.ToString();
                    }
                    else if (fileName.ToUpper().Contains("APPROVAL_EXTRACT"))
                    {
                        fileNames[1] = fileName.ToString();
                    }
                    else if (fileName.ToUpper().Contains("NA INTEGRATED DASHBOARD"))
                    {
                        fileNames[2] = fileName.ToString();
                    }
                    else if (fileName.ToUpper().Contains("PRODUCTION BUDGET"))
                    {
                        fileNames[3] = fileName.ToString();
                    }
                    else if (fileName.ToUpper().Contains("DAILY PRODUCTION DASHBOARD"))
                    {
                        fileNames[4] = fileName.ToString();
                    }
                    else ;
                }

            }

            return fileNames;
        }
       
    }
}

