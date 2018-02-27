using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using Word=Microsoft.Office.Interop.Word;

namespace CalculadoraDeTraduccionAustria
{
    public partial class MainForm : Form, IMainObserver
    {
        private WordDocumentFile document;
        private AddJobForm addJob;
        private int totalLineas;
        private Thread workerThread;
        private string selectedPath;
        private ProgressBarForm progressBarForm = new ProgressBarForm();
        private ImportSettings settings;
        private List<string> descriptionsList = new List<string>();
        private ComboBox cb;
        private string descriptionFirstElement;

        public MainForm()
        {
            InitializeComponent();
            settings = new ImportSettings();
            descriptionsList = settings.ReadDescriptionsSettins();
            labelTotalText.Text = "0,00";
       
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            cb = new ComboBox();
            foreach (string des in descriptionsList)
            {
                cb.Items.Add(des);
                descriptionFirstElement = des;
            }
        }

        public void UpdateElements()
        {
            if(progressBarForm.CancelTask)
            {
                this.workerThread.Abort();
            }
            else
            {
                //SetFileInfo(addJob.FileName, addJob.Description, addJob.Lines, addJob.Date);
                SetFileInfo(addJob.FileName, descriptionFirstElement, addJob.Lines, addJob.Date);             
            }           
        }

        #region Buttons
        /// <summary>
        /// Inicia un cuadro de dialogo para poder seleccionar un archivo word.
        /// De momento solo es posible agregar de un archivo a la vez
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonAddWork_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Word Document (*.docx, *.doc)| *.docx; *.doc";
            DialogResult result = openFileDialog1.ShowDialog();
            if(result == DialogResult.OK)
            {
                //document = new WordDocumentFile(openFileDialog1.FileName);
                //var date = File.GetLastWriteTime(openFileDialog1.FileName).ToString("dd/MM/yyyy");
                //SetFileInfo(document.DocumentName, descriptionFirstElement, document.DocumentCharactersCount, date);
                StarThreadDataGridViewFileSource(openFileDialog1.FileName);
            }
        }

        private void StarThreadDataGridViewFileSource(string fileName)
        {
            workerThread = new Thread(() => LoadFileFromWordAddMenu(fileName));
            workerThread.Start();
            StartProgresBarForm();
        }

        private void LoadFileFromWordAddMenu(string fileName)
        {
            document = new WordDocumentFile(openFileDialog1.FileName);
            var date = File.GetLastWriteTime(openFileDialog1.FileName).ToString("dd/MM/yyyy");
            SetFileInfo(document.DocumentName, descriptionFirstElement, document.DocumentCharactersCount, date);
            progressBarForm.ShowProgressBar = false;
            progressBarForm.UpdateProgressBar();
        }

        private void buttonSelectFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                selectedPath = folderBrowserDialog1.SelectedPath;
                StartThreadGridViewFolderSource();
            }
        }

        /// <summary>
        /// Inicia un cuadro de dialogo para agregar un archivo de cualquier tipo
        /// el cual no será procesado de forma automatica
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonAddPowerPoint_Click(object sender, EventArgs e)
        {
            addJob = new AddJobForm();
            addJob.RegisterObs(this);
            addJob.ShowDialog(this);
        }

        /// <summary>
        /// Elimina algun documento seleccionado en el 
        /// datagridview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRemoveWork_Click(object sender, EventArgs e)
        {
            RemoveDataSoruce();
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Word Documents (*.docx)|*.docx";
            saveFileDialog1.FileName = "export.docx";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //ToCsV(dataGridView1, @"c:\export.xls");
                //ToCsv(dataGridView1, saveFileDialog1.FileName); // Here dataGridview1 is your grid view name
                Export_Data_To_Word(dataGridView1, saveFileDialog1.FileName);
            }
        }
        #endregion Buttons


        public void SetFileInfo(string fileName, string description, int characters, string date)
        {
            var lineas = TotalLineas(Convert.ToInt32(textBoxSimbolosLineas.Text), characters);
            var precio = Convert.ToInt32(lineas) * Convert.ToDecimal(textBoxValorLinea.Text);
            TranslationInfo info = new TranslationInfo
            {
                Select = false,
                FileName = fileName,
                Description = description,
                Date = date,
                Lines = lineas,
                Precio = precio
            };
            SetDataSource(info);
        }


        /// <summary>
        /// Calcula la cantidad total de lineas
        /// tomando en cuenta la cantidad definida de simbolos por linea
        /// </summary>
        /// <param name="simbolosXlinea">Cantidad definida de simbolos por lineas</param>
        /// <param name="characters"></param>
        /// <returns></returns>
        public int TotalLineas(int simbolosXlinea, int characters)
        {
            var total = Decimal.Divide(characters, simbolosXlinea);
            totalLineas = Convert.ToInt32(Math.Ceiling(total));
            return totalLineas;
        }

        internal delegate void SetDataSourceDelegate(TranslationInfo info);


        public void SetDataSource(TranslationInfo info)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new SetDataSourceDelegate(SetDataSource), info);
            }
            else
            {
                ((DataGridViewComboBoxColumn)dataGridView1.Columns["Beschreibung"]).DataSource = cb.Items;
                //dataGridView1.Rows.Add(info.Select, info.FileName, info.Description, info.Date, info.Lines, info.Precio.ToString("N2"));
                //BindingList<string> bindingList = new BindingList<string>();               
                dataGridView1.Rows.Add(info.Select, info.FileName, descriptionFirstElement, info.Date, info.Lines, info.Precio.ToString("N2"));
                dataGridView1.Refresh();
                UpdatePrice();
            }
        }

        internal delegate void RemoveDataSoruceDelegate();

        public void RemoveDataSoruce()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells["Select"].Value) == true)
                {
                    dataGridView1.Rows.RemoveAt(row.Index);
                }
                else { }
            }
            UpdatePrice();
            dataGridView1.Refresh();
        }

        public void UpdatePrice()
        {
            Decimal total = 0;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                total += Convert.ToDecimal(row.Cells["Betrag"].Value);
            }
            labelTotalText.Text = total.ToString();
        }


        private void StartProgresBarForm()
        {
            progressBarForm.RegisterObs(this);
            progressBarForm.ShowDialog(this);

            progressBarForm.ShowProgressBar = true;
            progressBarForm.CancelTask = false;
        }


        private void StartThreadGridViewFolderSource()
        {
            workerThread = new Thread(new ThreadStart(LoadFilesFromFolder));
            workerThread.Start();
            StartProgresBarForm();
            //progressBarForm.ShowDialog(this);
        }

        internal delegate void LoadFilesFromFolderDelegate();

        /// <summary>
        /// Devuelve todos los archivos Word que se encuentran
        /// en la carpeta seleccionada
        /// </summary>
        /// <param name="path"></param>
        public void LoadFilesFromFolder()
        {
            var ext = new List<string> { ".docx", ".doc" };
            var myFiles = Directory.GetFiles(selectedPath, "*.*", SearchOption.AllDirectories).Where(s => ext.Contains(Path.GetExtension(s)));
            //var myFiles = Directory.GetFiles(selectedPath, "*.*", SearchOption.AllDirectories).Where(file => new string[] { ".docx", ".doc" }.Contains(Path.GetExtension(file)));
                List<string> Files = myFiles.OfType<string>().ToList();
                //var myFiles = Directory.GetFiles(path);
                foreach(string file in Files)
                {
                    document = new WordDocumentFile(file);
                    var date = File.GetLastWriteTime(file).ToString("dd/MM/yyyy");
                //SetFileInfo(document.DocumentName, "Ansprüche EN-DE", document.DocumentCharactersCount, date);
                SetFileInfo(document.DocumentName, descriptionFirstElement, document.DocumentCharactersCount, date);
            }
            progressBarForm.ShowProgressBar = false;
            progressBarForm.UpdateProgressBar();
        }

        private void ToCsv(DataGridView dGV, string filename)
        {
            string stOutput = "";
            // Export titles:
            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            // Export data.
            for (int i = 0; i < dGV.RowCount - 1; i++)
            {
                string stLine = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }


        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 12;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                //oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "Rechnung";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file
                oDoc.SaveAs2(filename);

            }
        }
    }
}
