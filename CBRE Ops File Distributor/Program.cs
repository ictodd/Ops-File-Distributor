using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CBRE_Ops_File_Distributor {

    public enum OfficeApp {
        Excel,
        Word
    }
    
    class Program {

        static bool DebugOn = false;

        static int CompletedTasks = 0;
        static int ErrorCount = 0;

        #region File and Folder Paths

        // contains the excel (Forbury) and word files
        private const string SOURCE_NEW_TEMPLATE_PATH_AKL = @"\\NZAKLFNP03\Departments\Valuations\OPERATIONS\Projects\New Template Project\Latest Model & Template";
        private const string SOURCE_FORBURY = SOURCE_NEW_TEMPLATE_PATH_AKL + @"\\CBRE Forbury Template.xlsm";
        private const string SOURCE_CMV_TEMPLATE = SOURCE_NEW_TEMPLATE_PATH_AKL + @"\\CMV Template.docm";
        private const string SOURCE_GROUND_RENT_CERT = SOURCE_NEW_TEMPLATE_PATH_AKL + @"\\Ground Rent Cert.docm";
        private const string SOURCE_GROUND_RENT_TEMPLATE = SOURCE_NEW_TEMPLATE_PATH_AKL + @"\\Ground Rent Template.docm";
        private const string SOURCE_RENT_REVIEW_TEMPLATE = SOURCE_NEW_TEMPLATE_PATH_AKL + @"\\RR Template.docm";
        private const string SOURCE_RENT_REVIEW_CERT = SOURCE_NEW_TEMPLATE_PATH_AKL + @"\\RR Cert.docm";

        private const string DESTIN_NEW_TEMPLATE_PATH_AKL = @"\\NZAKLFNP03\Departments\Valuations\_CBRE_NZVAS_Master_Models_Reports\New Report 2018";
        private const string DESTIN_NEW_TEMPLATE_PATH_AKL_CMV = DESTIN_NEW_TEMPLATE_PATH_AKL + @"\\CMV";
        private const string DESTIN_NEW_TEMPLATE_PATH_AKL_GROUND_RENT = DESTIN_NEW_TEMPLATE_PATH_AKL + @"\\Ground Rent";
        private const string DESTIN_NEW_TEMPLATE_PATH_AKL_RENT_REVIEW = DESTIN_NEW_TEMPLATE_PATH_AKL + @"\\Rent Review";

        // WELLINGTON

        private const string DESTIN_NEW_TEMPLATE_PATH_WLG = @"\\NZWLGFNP03\Departments\Valuations\TEMPLATES\Forbury_Model";

        // Wellington are being annoying and keep changing the folder structure. Just placing all docs into above directory

        //private const string DESTIN_NEW_TEMPLATE_PATH_WLG_CMV = DESTIN_NEW_TEMPLATE_PATH_WLG + @"\\CMV";
        //private const string DESTIN_NEW_TEMPLATE_PATH_WLG_GROUND_RENT = DESTIN_NEW_TEMPLATE_PATH_WLG + @"\\Ground Rent";
        //private const string DESTIN_NEW_TEMPLATE_PATH_WLG_RENT_REVIEW = DESTIN_NEW_TEMPLATE_PATH_WLG + @"\\Rent Review";

        // CHRISTCHURCH

        private const string DESTIN_NEW_TEMPLATE_PATH_CHC = @"\\NZCHCFNP03\Departments\Valuations\Template\New Report 2018";
        private const string DESTIN_NEW_TEMPLATE_PATH_CHC_CMV = DESTIN_NEW_TEMPLATE_PATH_CHC + @"\\CMV";
        private const string DESTIN_NEW_TEMPLATE_PATH_CHC_GROUND_RENT = DESTIN_NEW_TEMPLATE_PATH_CHC + @"\\Ground Rent";
        private const string DESTIN_NEW_TEMPLATE_PATH_CHC_RENT_REVIEW = DESTIN_NEW_TEMPLATE_PATH_CHC + @"\\Rent Review";

        // contains the insurance model and report that needs to be copied to wlg and chch
        private const string SOURCE_INS_TEMPLATE_PATH_AKL = @"\\NZAKLFNP03\Departments\Valuations\_CBRE_NZVAS_Master_Models_Reports\Insurance\Current Template";

        private const string DESTIN_INS_TEMPLATE_PATH_WLG = @"\\NZWLGFNP03\Departments\Valuations\TEMPLATES\Insurance_Val_XXXX_INS\New Template";
        private const string DESTIN_INS_TEMPLATE_PATH_CHC = @"\\NZCHCFNP03\Departments\Valuations\Template\Insurance";

        // cbre add ins (just main word and excel ones)
        private const string SOURCE_CBRE_VAL_TOOLS_EXCEL = @"\\NZAKLFNP03\Departments\Valuations\TODDS APP STORE\Add-Ins\Excel\CBRE Val Tools.xlam";
        private const string SOURCE_CBRE_VAL_TOOLS_WORD = @"\\NZAKLFNP03\Departments\Valuations\TODDS APP STORE\Add-Ins\Word\CBRE Val Tools - Word.dotm";

        private const string DESTIN_TOOLS_EXCEL_WLG = @"\\NZWLGFNP03\Departments\Valuations\TODDS APP STORE\Add-Ins\Excel";
        private const string DESTIN_TOOLS_WORD_WLG = @"\\NZWLGFNP03\Departments\Valuations\TODDS APP STORE\Add-Ins\Word";

        private const string DESTIN_TOOLS_EXCEL_CHC = @"\\NZCHCFNP03\Departments\Valuations\TODDS APP STORE\Add-Ins\Excel";
        private const string DESTIN_TOOLS_WORD_CHC = @"\\NZCHCFNP03\Departments\Valuations\TODDS APP STORE\Add-Ins\Word";

        #endregion

        static void Main(string[] args) {

            Menu.Header();
            Menu.CheckPassword();
            // Menu.DisplayMain();
            Menu.Continue();
                        
            // Get ins templates
            string insExcelPath = GetFile(OfficeApp.Excel, SOURCE_INS_TEMPLATE_PATH_AKL);
            string insWordPath = GetFile(OfficeApp.Word, SOURCE_INS_TEMPLATE_PATH_AKL);

            Console.WriteLine();

            List<Task> allTasks = new List<Task>();
            Task t;

            // MODELS AND TEMPLATES

            // Auckland
            t = Task.Run(() => CopyTemplateFile(SOURCE_FORBURY, DESTIN_NEW_TEMPLATE_PATH_AKL));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_CMV_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_AKL_CMV));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_GROUND_RENT_CERT, DESTIN_NEW_TEMPLATE_PATH_AKL_GROUND_RENT));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_GROUND_RENT_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_AKL_GROUND_RENT));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_RENT_REVIEW_CERT, DESTIN_NEW_TEMPLATE_PATH_AKL_RENT_REVIEW));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_RENT_REVIEW_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_AKL_RENT_REVIEW));
            allTasks.Add(t);

            // Wellington
            t = Task.Run(() => CopyTemplateFile(SOURCE_FORBURY, DESTIN_NEW_TEMPLATE_PATH_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_CMV_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_GROUND_RENT_CERT, DESTIN_NEW_TEMPLATE_PATH_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_GROUND_RENT_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_RENT_REVIEW_CERT, DESTIN_NEW_TEMPLATE_PATH_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_RENT_REVIEW_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_WLG));
            allTasks.Add(t);

            // Christchurch
            t = Task.Run(() => CopyTemplateFile(SOURCE_FORBURY, DESTIN_NEW_TEMPLATE_PATH_CHC));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_CMV_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_CHC_CMV));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_GROUND_RENT_CERT, DESTIN_NEW_TEMPLATE_PATH_CHC_GROUND_RENT));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_GROUND_RENT_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_CHC_GROUND_RENT));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_RENT_REVIEW_CERT, DESTIN_NEW_TEMPLATE_PATH_CHC_RENT_REVIEW));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_RENT_REVIEW_TEMPLATE, DESTIN_NEW_TEMPLATE_PATH_CHC_RENT_REVIEW));
            allTasks.Add(t);

            // INSURANCE
            t = Task.Run(() => CopyTemplateFile(insWordPath, DESTIN_INS_TEMPLATE_PATH_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(insWordPath, DESTIN_INS_TEMPLATE_PATH_CHC));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(insExcelPath, DESTIN_INS_TEMPLATE_PATH_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(insExcelPath, DESTIN_INS_TEMPLATE_PATH_CHC));
            allTasks.Add(t);

            // CBRE ADD INS
            t = Task.Run(() => CopyTemplateFile(SOURCE_CBRE_VAL_TOOLS_EXCEL, DESTIN_TOOLS_EXCEL_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_CBRE_VAL_TOOLS_EXCEL, DESTIN_TOOLS_EXCEL_CHC));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_CBRE_VAL_TOOLS_WORD, DESTIN_TOOLS_WORD_WLG));
            allTasks.Add(t);

            t = Task.Run(() => CopyTemplateFile(SOURCE_CBRE_VAL_TOOLS_WORD, DESTIN_TOOLS_WORD_CHC));
            allTasks.Add(t);

            Task.WaitAll(allTasks.ToArray());

            Console.WriteLine($"Finished a total of {CompletedTasks} with a total of {ErrorCount} errors.");

            // Finish up
            Menu.Finish();

        }

        private static bool CopyTemplateFile(string sourceFile, string destinationRootPath) {
            bool result = false;
            
            if (sourceFile != null) {
                string destinationFile = destinationRootPath + "\\" + Path.GetFileName(sourceFile);
                try {
                    File.Copy(sourceFile, destinationFile,true);
                    CompletedTasks++;
                    Console.WriteLine("Completed Task: copied '" + Path.GetFileName(sourceFile) + "' from " + Path.GetDirectoryName(sourceFile) + "' to " + Path.GetDirectoryName(destinationFile) + "\n");
                    result = true;
                } catch(Exception ex) {
                    Console.WriteLine("Failed to copy '" + Path.GetFileName(sourceFile) +
                                        "' from '" + Path.GetDirectoryName(sourceFile) + "' to " +
                                        Path.GetDirectoryName(destinationFile));
                    Console.WriteLine("\tError Message: " + ex.Message + "\n");
                    ErrorCount++;
                }
            } else {
                Console.WriteLine("There was an issue getting " + Path.GetFileName(sourceFile) + " from " + Path.GetDirectoryName(sourceFile) + " source file destined for '" + destinationRootPath + "'\n");
            }
            return result;
        }

        private static bool ValidExtension(string filePath, OfficeApp app) {
            
            // exclude temp files
            if (Path.GetFileName(filePath).Contains("~$"))
                return false;

            string extn = Path.GetExtension(filePath);
            
            if (app == OfficeApp.Excel) {
                switch (extn) {
                    case ".xlsm":
                    case ".xlsx":
                    case ".xls":
                    case ".xlsb":
                        return true;
                }
            }else if(app == OfficeApp.Word) {
                switch (extn) {
                    case ".doc":
                    case ".docm":
                    case ".docx":
                        return true;
                }
            }

            return false;
        }

        private static string GetFile(OfficeApp fileType, string path) {

            DateTime latestDate = new DateTime();
            string result = null;

            foreach(string file in Directory.GetFiles(path)) {

                FileInfo fi = new FileInfo(file);

                //DEBUG_PRINT("Checking file " + file);
                //DEBUG_PRINT("Has a creation time of " + fi.CreationTime.ToShortDateString());
                //DEBUG_PRINT("Current latest date is " + latestDate.ToShortDateString());

                if (ValidExtension(file,fileType) && fi.CreationTime > latestDate) {

                    result = file;
                    latestDate = fi.CreationTime;

                }

            }

            DEBUG_PRINT("File from " + path + " of type " + fileType.ToString() + " with a result of " + result);

            return result;
        }

        private static void DEBUG_PRINT(string message) {
            if (DebugOn) Console.WriteLine("**| DEBUG MESSAGE: " + message + " |**\n");
        }
    }
}
