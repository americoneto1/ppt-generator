using System;
using System.Configuration;
using System.Windows;

namespace LouvorPPT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            txtTemplate.Text = ConfigurationManager.AppSettings["defaultTemplatePath"].ToString();
            txtDestination.Text = ConfigurationManager.AppSettings["defaultDestinationPath"].ToString();
        }

        private void btnGerar_Click(object sender, RoutedEventArgs e)
        {
            if (CheckRequiredFields())
            {
                MessageBox.Show("Preencha todos os campos antes de prosseguir", "Preenchimento obrigatório", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            try
            {
                Configuracao objConfig = new Configuracao();
                objConfig.TemplateFile = txtTemplate.Text;
                objConfig.DestinationPath = txtDestination.Text;

                new Presentation(objConfig).Generate(txtTitle.Text, txtContent.Text);

                MessageBox.Show(string.Format("A apresentação foi gerada no caminho {0}!", objConfig.DestinationPath), "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        private bool CheckRequiredFields()
        {
            return string.IsNullOrEmpty(txtTemplate.Text) || string.IsNullOrEmpty(txtDestination.Text) ||
                string.IsNullOrEmpty(txtContent.Text) || string.IsNullOrEmpty(txtTitle.Text);
        }
    }
}
