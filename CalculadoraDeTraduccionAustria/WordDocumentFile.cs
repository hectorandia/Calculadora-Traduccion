﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace CalculadoraDeTraduccionAustria
{
    public class WordDocumentFile
    {
        private string documentName;
        private int documentCharactersCount;
        private int totalLineas;
        private Application app;
        private Document document;

        public WordDocumentFile(string documentPath)
        {
            //Application ap = new Application();
            app = new Application();
            document = app.Documents.Open(documentPath, Type.Missing, true);
            Range rng = document.Content;
            rng.Select();

            DocumentCharactersCount = rng.ComputeStatistics(WdStatistic.wdStatisticCharactersWithSpaces);
            DocumentName = document.Name;
            document.Close();
            //CloseWord();
            ap.Quit(false);
        }

        public void CloseWord()
        {
            
        }

        public string DocumentName
        {
            get
            {
                return documentName;
            }
            set
            {
                documentName = value;
            }
        }

        public int DocumentCharactersCount
        {
            get
            {
                return documentCharactersCount;
            }
            set
            {
                documentCharactersCount = value;
            }

        }

        public int TotalLineas(int simbolosXlinea)
        {
            var total = Decimal.Divide(DocumentCharactersCount, simbolosXlinea);
            totalLineas = Convert.ToInt32(Math.Ceiling(total));
            return totalLineas;
        }
    }
}
