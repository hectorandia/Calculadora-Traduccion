using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data;

namespace CalculadoraDeTraduccionAustria
{
    public class ImportSettings
    {
        //private string[] DescriptionsList;

        public List<string> ReadDescriptionsSettins()
        {
            List<string> list = new List<string>();
            string[] DescriptionsList = ConfigurationManager.AppSettings["Descriptions"].Split(',');
            list = DescriptionsList.OfType<string>().ToList();
            
            return list;
        }
    }

}
