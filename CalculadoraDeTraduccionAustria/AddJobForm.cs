using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalculadoraDeTraduccionAustria
{
    public partial class AddJobForm : Form, ISubject
    {
        public string FileName { get; set; }
        public string Description { get; set; }
        public int Lines { get; set; }
        public string Date { get; set; }
        private List<IMainObserver> observers = new List<IMainObserver>();
        private List<string> descriptionsList = new List<string>();
        ImportSettings settings;

        public AddJobForm()
        {
            InitializeComponent();
            settings = new ImportSettings();
            descriptionsList = settings.ReadDescriptionsSettins();
            foreach(string des in descriptionsList)
            {
                comboBoxDescription.Items.Add(des);
                comboBoxDescription.SelectedItem = des;
            }
            SetAlertTextOff();
            
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            if(ValidateContent())
            {
                FileName = textBoxFileName.Text;
                Description = comboBoxDescription.Text;
                Lines = Convert.ToInt32(textBoxLines.Text);
                Date = monthCalendar1.SelectionRange.Start.ToShortDateString();

                NotifyObs();
                this.Close();
                this.Dispose();
            }      
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
        }

        public bool ValidateContent()
        {
            bool valid = true;
            SetAlertTextOff();
            if (String.IsNullOrEmpty(this.textBoxFileName.Text))
            {
                valid = false;
                labelFileNameAlert.Visible = true;
            }

            if(String.IsNullOrEmpty(this.textBoxLines.Text))
            {
                valid = false;
                labelLinesAlert.Visible = true;
            }

            return valid;
        }

        public void SetAlertTextOff()
        {
            labelFileNameAlert.Visible = false;
            labelLinesAlert.Visible = false;
        }

        public void RegisterObs(IMainObserver ob)
        {
            observers.Add(ob);
        }

        public void UnregisterObs(IMainObserver ob)
        {
            observers.Remove(ob);
        }

        public void NotifyObs()
        {
            foreach (IMainObserver ob in observers)
            {
                ob.UpdateElements();
            }
        }

        private void closePanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics z = e.Graphics;
            Pen pen = new Pen(Color.FromArgb(169, 169, 169));
            z.DrawLine(pen, 3, 3, 15, 15);
            z.DrawLine(pen, 3, 15, 15, 3);
            z.DrawLine(pen, 4, 3, 16, 15);
            z.DrawLine(pen, 4, 15, 16, 3);
        }
    }
}
