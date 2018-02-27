using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalculadoraDeTraduccionAustria
{
    public interface ISubject
    {
        void RegisterObs(IMainObserver ob);
        void UnregisterObs(IMainObserver ob);
        void NotifyObs();
    }
}
