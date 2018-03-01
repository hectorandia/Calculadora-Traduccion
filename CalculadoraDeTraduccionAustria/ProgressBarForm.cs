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
    public partial class ProgressBarForm : Form, ISubject
    {
        public bool ShowProgressBar { get; set; }
        public bool CancelTask { get; set; }
        private List<IMainObserver> observers = new List<IMainObserver>();

        public ProgressBarForm()
        {
            InitializeComponent();
            progressBar1.Style = ProgressBarStyle.Marquee;
            CheckForIllegalCrossThreadCalls = false;
            CancelTask = false;
        }

        public void UpdateProgressBar()
        {
            if (ShowProgressBar)
            {

            }
            else
            {
                //this.Close();
                NotifyObs();
            }
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

        private void ProgressBarForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            CancelTask = true;
            NotifyObs();
        }
    }
}
