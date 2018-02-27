namespace CalculadoraDeTraduccionAustria
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Select = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Übersetzung = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Beschreibung = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.Ausführungsdatum = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Normzeilen = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Betrag = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1 = new System.Windows.Forms.Panel();
            this.buttonSave = new System.Windows.Forms.Button();
            this.labelTotalText = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.buttonAddPowerPoint = new System.Windows.Forms.Button();
            this.buttonSelectFolder = new System.Windows.Forms.Button();
            this.buttonRemoveWork = new System.Windows.Forms.Button();
            this.buttonAddWork = new System.Windows.Forms.Button();
            this.textBoxSimbolosLineas = new System.Windows.Forms.TextBox();
            this.textBoxValorLinea = new System.Windows.Forms.TextBox();
            this.labelSimboloLinea = new System.Windows.Forms.Label();
            this.labelValorLinea = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Select,
            this.Übersetzung,
            this.Beschreibung,
            this.Ausführungsdatum,
            this.Normzeilen,
            this.Betrag});
            this.dataGridView1.Location = new System.Drawing.Point(12, 130);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1043, 587);
            this.dataGridView1.TabIndex = 0;
            // 
            // Select
            // 
            this.Select.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Select.HeaderText = "Select";
            this.Select.Name = "Select";
            this.Select.Width = 43;
            // 
            // Übersetzung
            // 
            this.Übersetzung.HeaderText = "Übersetzung";
            this.Übersetzung.MinimumWidth = 10;
            this.Übersetzung.Name = "Übersetzung";
            // 
            // Beschreibung
            // 
            this.Beschreibung.HeaderText = "Beschreibung";
            this.Beschreibung.MinimumWidth = 100;
            this.Beschreibung.Name = "Beschreibung";
            this.Beschreibung.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Beschreibung.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // Ausführungsdatum
            // 
            this.Ausführungsdatum.HeaderText = "Ausführungsdatum";
            this.Ausführungsdatum.Name = "Ausführungsdatum";
            // 
            // Normzeilen
            // 
            this.Normzeilen.HeaderText = "Normzeilen";
            this.Normzeilen.Name = "Normzeilen";
            this.Normzeilen.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // Betrag
            // 
            this.Betrag.HeaderText = "Betrag";
            this.Betrag.MinimumWidth = 100;
            this.Betrag.Name = "Betrag";
            this.Betrag.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.buttonSave);
            this.panel1.Controls.Add(this.labelTotalText);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 723);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1067, 125);
            this.panel1.TabIndex = 1;
            // 
            // buttonSave
            // 
            this.buttonSave.Location = new System.Drawing.Point(25, 22);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(92, 23);
            this.buttonSave.TabIndex = 2;
            this.buttonSave.Text = "Save";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // labelTotalText
            // 
            this.labelTotalText.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelTotalText.AutoSize = true;
            this.labelTotalText.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTotalText.Location = new System.Drawing.Point(953, 30);
            this.labelTotalText.Name = "labelTotalText";
            this.labelTotalText.Size = new System.Drawing.Size(36, 16);
            this.labelTotalText.TabIndex = 1;
            this.labelTotalText.Text = "0,00";
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(843, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Gesamtbetrag:";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.buttonAddPowerPoint);
            this.panel2.Controls.Add(this.buttonSelectFolder);
            this.panel2.Controls.Add(this.buttonRemoveWork);
            this.panel2.Controls.Add(this.buttonAddWork);
            this.panel2.Controls.Add(this.textBoxSimbolosLineas);
            this.panel2.Controls.Add(this.textBoxValorLinea);
            this.panel2.Controls.Add(this.labelSimboloLinea);
            this.panel2.Controls.Add(this.labelValorLinea);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 24);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1067, 100);
            this.panel2.TabIndex = 2;
            // 
            // buttonAddPowerPoint
            // 
            this.buttonAddPowerPoint.Location = new System.Drawing.Point(923, 25);
            this.buttonAddPowerPoint.Name = "buttonAddPowerPoint";
            this.buttonAddPowerPoint.Size = new System.Drawing.Size(111, 23);
            this.buttonAddPowerPoint.TabIndex = 7;
            this.buttonAddPowerPoint.Text = "Add Power Point";
            this.buttonAddPowerPoint.UseVisualStyleBackColor = true;
            this.buttonAddPowerPoint.Click += new System.EventHandler(this.buttonAddPowerPoint_Click);
            // 
            // buttonSelectFolder
            // 
            this.buttonSelectFolder.Location = new System.Drawing.Point(689, 25);
            this.buttonSelectFolder.Name = "buttonSelectFolder";
            this.buttonSelectFolder.Size = new System.Drawing.Size(111, 23);
            this.buttonSelectFolder.TabIndex = 6;
            this.buttonSelectFolder.Text = "Select Folder";
            this.buttonSelectFolder.UseVisualStyleBackColor = true;
            this.buttonSelectFolder.Click += new System.EventHandler(this.buttonSelectFolder_Click);
            // 
            // buttonRemoveWork
            // 
            this.buttonRemoveWork.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonRemoveWork.Location = new System.Drawing.Point(923, 54);
            this.buttonRemoveWork.Name = "buttonRemoveWork";
            this.buttonRemoveWork.Size = new System.Drawing.Size(111, 23);
            this.buttonRemoveWork.TabIndex = 5;
            this.buttonRemoveWork.Text = "Delete File";
            this.buttonRemoveWork.UseVisualStyleBackColor = true;
            this.buttonRemoveWork.Click += new System.EventHandler(this.buttonRemoveWork_Click);
            // 
            // buttonAddWork
            // 
            this.buttonAddWork.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonAddWork.Location = new System.Drawing.Point(806, 25);
            this.buttonAddWork.Name = "buttonAddWork";
            this.buttonAddWork.Size = new System.Drawing.Size(111, 23);
            this.buttonAddWork.TabIndex = 4;
            this.buttonAddWork.Text = "Add File";
            this.buttonAddWork.UseVisualStyleBackColor = true;
            this.buttonAddWork.Click += new System.EventHandler(this.buttonAddWork_Click);
            // 
            // textBoxSimbolosLineas
            // 
            this.textBoxSimbolosLineas.Location = new System.Drawing.Point(127, 54);
            this.textBoxSimbolosLineas.Name = "textBoxSimbolosLineas";
            this.textBoxSimbolosLineas.Size = new System.Drawing.Size(123, 20);
            this.textBoxSimbolosLineas.TabIndex = 3;
            this.textBoxSimbolosLineas.Text = "55";
            // 
            // textBoxValorLinea
            // 
            this.textBoxValorLinea.Location = new System.Drawing.Point(127, 20);
            this.textBoxValorLinea.Name = "textBoxValorLinea";
            this.textBoxValorLinea.Size = new System.Drawing.Size(123, 20);
            this.textBoxValorLinea.TabIndex = 2;
            this.textBoxValorLinea.Text = "0,7";
            // 
            // labelSimboloLinea
            // 
            this.labelSimboloLinea.AutoSize = true;
            this.labelSimboloLinea.Location = new System.Drawing.Point(22, 58);
            this.labelSimboloLinea.Name = "labelSimboloLinea";
            this.labelSimboloLinea.Size = new System.Drawing.Size(95, 13);
            this.labelSimboloLinea.TabIndex = 1;
            this.labelSimboloLinea.Text = "Simbolos por linea:";
            // 
            // labelValorLinea
            // 
            this.labelValorLinea.AutoSize = true;
            this.labelValorLinea.Location = new System.Drawing.Point(22, 23);
            this.labelValorLinea.Name = "labelValorLinea";
            this.labelValorLinea.Size = new System.Drawing.Size(77, 13);
            this.labelValorLinea.TabIndex = 0;
            this.labelValorLinea.Text = "Valor por linea:";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1067, 24);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 848);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(600, 450);
            this.Name = "MainForm";
            this.Text = "Media-Translations";
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.Label labelValorLinea;
        private System.Windows.Forms.Label labelSimboloLinea;
        private System.Windows.Forms.TextBox textBoxSimbolosLineas;
        private System.Windows.Forms.TextBox textBoxValorLinea;
        private System.Windows.Forms.Button buttonAddWork;
        private System.Windows.Forms.Button buttonRemoveWork;
        private System.Windows.Forms.Label labelTotalText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button buttonSelectFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button buttonAddPowerPoint;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Select;
        private System.Windows.Forms.DataGridViewTextBoxColumn Übersetzung;
        private System.Windows.Forms.DataGridViewComboBoxColumn Beschreibung;
        private System.Windows.Forms.DataGridViewTextBoxColumn Ausführungsdatum;
        private System.Windows.Forms.DataGridViewTextBoxColumn Normzeilen;
        private System.Windows.Forms.DataGridViewTextBoxColumn Betrag;
    }
}