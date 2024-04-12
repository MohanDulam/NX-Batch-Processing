using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace NX_Batch_Processing
{
    public partial class NX_Batch_Processing : Form
    {
        public NX_Batch_Processing()
        {
            InitializeComponent();
        }

        private void btnBrowseInputFile_Click(object sender, EventArgs e)
        {
            txtInputFolder.Clear();
            lblInfo.Text = "Status Info...";
            lblInfo.ForeColor = Color.White;

            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                DialogResult dialogResult = fbd.ShowDialog();

                if (dialogResult == DialogResult.OK)

                    txtInputFolder.Text = fbd.SelectedPath.Trim();
            }
        }

        private void btnBrowseExportFolder_Click(object sender, EventArgs e)
        {
            txtExportFolder.Clear();

            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                DialogResult dialogResult = fbd.ShowDialog();

                if (dialogResult == DialogResult.OK)

                    txtExportFolder.Text = fbd.SelectedPath.Trim();
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            // Clear the text box content  
            txtInputFolder.Clear();
            txtExportFolder.Clear();

            // Uncheck the all the Check boxes
            chkbxStep.Checked = false;
            chkbxIGES.Checked = false;
            chkbxParasolid.Checked = false;
            chkbxDWG.Checked = false;
            chkbxDXF.Checked = false;
            chkbxPDF.Checked = false;

            // Change to Status Info...
            lblInfo.Text = "Status Info...";
            lblInfo.ForeColor = Color.White;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                // Check Directory for both Part Files and Export Files
                if (string.IsNullOrWhiteSpace(txtInputFolder.Text) &&
                    string.IsNullOrWhiteSpace(txtExportFolder.Text))
                {
                    MessageBox.Show("Please Browse the Input Part File Directory and Export Files Directory");
                    return;
                }

                // Check for Tick option for different Export Options 
                if (chkbxStep.Checked == false && chkbxIGES.Checked == false && chkbxParasolid.Checked == false
                    && chkbxDWG.Checked == false && chkbxDXF.Checked == false && chkbxPDF.Checked == false)
                {
                    MessageBox.Show("Please any one of the Export Option");
                    return;
                }

                // Check for STEP File Export Option
                if (chkbxStep.Checked)
                {
                    lblInfo.Text = "Step Exporting is Processing";
                    lblInfo.ForeColor = Color.Green;
                    Program.ExportFiles(txtInputFolder.Text, "STEP-Files", "STEP", txtExportFolder.Text);
                    //Program.CreateStepFile(txtInputFolder.Text, txtExportFolder.Text);

                    lblInfo.Text = "Step Export Completed";

                    //MessageBox.Show("Export Part Files to Step Files");
                }

                // Check for IGES File Export Option
                if (chkbxIGES.Checked)
                {
                    lblInfo.Text = "IGES Exporting is Processing";
                    lblInfo.ForeColor = Color.Green;
                    //Program.CreateIGESFile(txtInputFolder.Text, txtExportFolder.Text);

                    Program.ExportFiles(txtInputFolder.Text, "IGES-Files", "IGES", txtExportFolder.Text);
                    lblInfo.Text = "IGES Export Completed";
                    //MessageBox.Show("ExportPart Files to IGES Fiels");
                }

                // Check for Parasolid File Export Option
                if (chkbxParasolid.Checked)
                {
                    lblInfo.Text = "Parasolid Exporting is Processing";
                    lblInfo.ForeColor = Color.Green;
                    //Program.CreateParasolidFile(txtInputFolder.Text, txtExportFolder.Text);

                    Program.ExportFiles(txtInputFolder.Text, "Parasolid-Files", "XT", txtExportFolder.Text);
                    lblInfo.Text = "Parasolid Export Completed";
                    //MessageBox.Show("ExportPart Files to IGES Fiels");
                }

                // Check for DWG File Export Option
                if (chkbxDWG.Checked)
                {
                    lblInfo.Text = "Drawing File Exporting is Processing";
                    lblInfo.ForeColor = Color.Green;

                    Program.ExportFiles(txtInputFolder.Text, "DWG-Files", "DWG", txtExportFolder.Text);

                    lblInfo.Text = "Drawing File is Export Completed";
                    //MessageBox.Show("DWG Export Coding is not Developed. under process ");
                }

                // Check for DXF File Export Option
                if (chkbxDXF.Checked)
                {
                    lblInfo.Text = "DXF File Exporting is Processing";
                    lblInfo.ForeColor = Color.Green;

                    Program.ExportFiles(txtInputFolder.Text, "DXF-Files", "DXF", txtExportFolder.Text);

                    lblInfo.Text = "DXF File is  Export Completed";
                    //MessageBox.Show("DXF Export Coding is not Developed. under process ");
                }

                // Check for PDF File Export Option
                if (chkbxPDF.Checked)
                {
                    lblInfo.Text = "PDF File Exporting is Processing";
                    lblInfo.ForeColor = Color.Green;

                    Program.ExportFiles(txtInputFolder.Text, "PDF-Files", "PDF", txtExportFolder.Text);

                    lblInfo.Text = "PDF File is Export Completed";
                    //MessageBox.Show("PDF Export Coding is not Developed. under process ");
                }

                MessageBox.Show("Export of files are Successfully Completed");
                //Program.TestMethod();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in Export files", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }
        }
    }
}

namespace NX_Batch_Processing
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new NX_Batch_Processing());
        }

        /// <summary>
        /// Export Part File to Required Format
        /// </summary>
        /// <param name="inputPartFiles">Part File with Full Path</param>
        /// <param name="folderName">Export Folder Name</param>
        /// <param name="fileExtension">Export File Extension</param>
        /// <param name="exportFilesDir">Export File Directory</param>
        public static void ExportFiles(string inputPartFiles, string folderName, string fileExtension, string exportFilesDir)
        {
            // Part Files and Export Files Directory
            string filesDirectoryPath = inputPartFiles;
            string[] files = Directory.GetFiles(filesDirectoryPath, "*.prt");
            string exportFolder = exportFilesDir + @"\" + folderName;
            string[] directories = Directory.GetDirectories(filesDirectoryPath);

            // Folder for Step Files
            Directory.CreateDirectory(exportFolder);

            // Declerartion Exe File Path
            string exeFileLoc;

            // Exe file loation
            // Get the file and path of the BatchExport.exe
            exeFileLoc = System.Reflection.Assembly.GetExecutingAssembly().Location;
            exeFileLoc = exeFileLoc.Substring(0, exeFileLoc.LastIndexOf("\\"));
            exeFileLoc = exeFileLoc + @"\BatchExport.exe";

            // Get the file and path of the run_managed.exe
            string argFirstSection = Environment.GetEnvironmentVariable("UGII_BASE_DIR") + @"\NXBIN\run_managed.exe";

            // output and error messages from ug_edit_part_names.exe
            string output, errorout;

            foreach (string file in files)
            {
                Console.WriteLine(file);
                string argSecondSection = $" \"{exeFileLoc}\"" + $" \"{file}\"" + $" \"{fileExtension}\"" + $" \"{exportFolder}\"";

                NewProcessor(argFirstSection, argSecondSection, out output, out errorout);

                argSecondSection = null;
                output = null; errorout = null;
            }
        }

        /// <summary>
        /// To Start New Processor 
        /// </summary>
        /// <param name="argFirstSectionOfAppllication">Appliction with Path</param>
        /// <param name="argSecondSection">Command for the Application</param>
        /// <param name="outputOfApplication">Output of Application</param>
        /// <param name="errorOfApplication">Errror of Application</param>
        private static void NewProcessor(string argFirstSectionOfAppllication, string argSecondSection, out string outputOfApplication, out string errorOfApplication)
        {
            outputOfApplication = null;
            errorOfApplication = null;
            try
            {
                // Arguments for the Processor 
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = argFirstSectionOfAppllication;
                startInfo.Arguments = argSecondSection;
                startInfo.RedirectStandardInput = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.UseShellExecute = false;

                // Intillazation of the New Processor
                Process process = new Process();
                process.StartInfo = startInfo;
                process.Start(); // Start Processor

                // Output and error from the Application
                outputOfApplication = process.StandardOutput.ReadToEnd();
                errorOfApplication = process.StandardError.ReadToEnd();

                process.WaitForExit();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in handling the Processor", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //throw;
            }
        }
    }
}
