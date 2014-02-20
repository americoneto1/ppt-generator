using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LouvorPPT
{
    public class Configuracao
    {
        private string _TemplateFile;

        public string TemplateFile
        {
            get { return _TemplateFile; }
            set {
                if (!System.IO.File.Exists(value))
                    throw new ArgumentException(string.Format("O arquivo de template {0} não existe", value));
                _TemplateFile = value; 
            }
        }

        private string _DestinationPath;

        public string DestinationPath
        {
            get { return _DestinationPath; }
            set {
                if (!System.IO.Directory.Exists(value))
                    throw new ArgumentException(string.Format("A pasta de destino {0} não existe", value));
                _DestinationPath = value; 
            }
        }
    }
}
