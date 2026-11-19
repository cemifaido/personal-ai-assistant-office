using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace PersonalAIAssistant
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Le continuazioni "await" devono tornare sul thread dell'interfaccia,
            // l'unico da cui si puo' usare il modello a oggetti di Outlook.
            if (!(SynchronizationContext.Current is WindowsFormsSynchronizationContext))
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            // Nient'altro all'avvio: Outlook disattiva gli add-in che rallentano l'apertura.
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region Codice generato da VSTO

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
