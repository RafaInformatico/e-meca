namespace EMecaAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        // Programming VSTO Add-Ins: https://msdn.microsoft.com/en-us/library/bb157876.aspx
        // Working with Text in Documents: https://msdn.microsoft.com/en-us/library/bb386277.aspx
        
        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
