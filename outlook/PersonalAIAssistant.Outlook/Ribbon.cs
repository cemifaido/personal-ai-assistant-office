using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace PersonalAIAssistant
{
    /// <summary>Gruppo "Personal AI" nella scheda Home della finestra principale di Outlook.</summary>
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private const string RibbonXml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab idMso='TabMail'>
        <group id='PaaGroup' label='Personal AI'>
          <button id='PaaSummarize' label='Riassumi selezionate' size='large' imageMso='ReadingPaneBottom'
                  screentip='Riassumi le email selezionate' supertip='Crea una bozza con il riassunto delle email selezionate (Ctrl+clic per selezionarne piu di una).'
                  onAction='OnSummarize' />
          <menu id='PaaReplyMenu' label='Rispondi con AI' size='large' imageMso='Reply'>
            <button id='PaaReplyConfirm' label='Conferma e ringrazia' onAction='OnReplyConfirm' />
            <button id='PaaReplyDecline' label='Declina gentilmente' onAction='OnReplyDecline' />
            <button id='PaaReplyClarify' label='Chiedi chiarimenti' onAction='OnReplyClarify' />
            <menuSeparator id='PaaSep' />
            <button id='PaaReplyCustom' label='Risposta personalizzata...' onAction='OnReplyCustom' />
          </menu>
          <button id='PaaSettings' label='Impostazioni' size='normal' imageMso='PropertySheet' onAction='OnSettings' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public string GetCustomUI(string ribbonID)
        {
            return ribbonID == "Microsoft.Outlook.Explorer" ? RibbonXml : null;
        }

        private static Microsoft.Office.Interop.Outlook.Application App { get { return Globals.ThisAddIn.Application; } }

        public void OnSummarize(Office.IRibbonControl c) { MailActions.SummarizeSelection(App); }
        public void OnReplyConfirm(Office.IRibbonControl c) { MailActions.Reply(App, "Ringrazia cordialmente e conferma l'accettazione."); }
        public void OnReplyDecline(Office.IRibbonControl c) { MailActions.Reply(App, "Declina gentilmente, ringraziando per l'opportunita' e indicando indisponibilita'."); }
        public void OnReplyClarify(Office.IRibbonControl c) { MailActions.Reply(App, "Chiedi maggiori dettagli, chiarimenti e tempistiche."); }
        public void OnReplyCustom(Office.IRibbonControl c) { MailActions.ReplyCustom(App); }

        public void OnSettings(Office.IRibbonControl c)
        {
            var s = AppSettings.Load();
            string intro = null;
            if (s.RecoveredFrom != null)
                intro = "Il file delle impostazioni mancava o era danneggiato: le ho recuperate dalla " + s.RecoveredFrom +
                        ". Controlla che siano corrette e premi Salva.";
            else if (s.ImportSaveFailed)
                intro = "Attenzione: non sono riuscito a salvare la copia cifrata delle chiavi della vecchia macro, " +
                        "quindi le ho lasciate nel registro (in chiaro) per non perderle. Reinseriscile qui e premi Salva: " +
                        "al prossimo avvio le rimuovero' dal registro.";
            else if (s.LegacyKeysRemain)
                intro = "Attenzione: nel registro restano chiavi della vecchia macro salvate in chiaro. Le rimuovo solo quando " +
                        "la stessa chiave risulta salvata qui (cifrata): se hai cambiato chiave, o il messaggio resta, cancella " +
                        @"la chiave di registro HKCU\Software\VB and VBA Program Settings\PersonalAIAssistant.";
            else if (s.ImportedFromMacro)
                intro = "Ho importato provider e chiavi dalla vecchia macro (ora sono cifrate) e le ho rimosse dal registro.";
            using (var f = new SettingsForm(s, intro)) f.ShowDialog();
        }
    }
}
