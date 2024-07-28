using System;
using Microsoft.Office.Core;
using System.IO;
using PPTCmd.Properties;

namespace PPTCmd
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            LoadAddin();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PPTCmd();
        }

        private void LoadAddin()
        {
            String AppDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            String addinPath = AppDataDir+"\\Microsoft\\AddIns";
            var macroFilePath = Path.Combine(addinPath, "MCMD.ppam");
            var addins = Globals.ThisAddIn.Application.AddIns.Add(macroFilePath);
            if (!(addins.Registered == MsoTriState.msoTrue && addins.Loaded == MsoTriState.msoTrue))
            {
                addins.Registered = MsoTriState.msoTrue;
                addins.Loaded = MsoTriState.msoTrue;
            }
        }
        public void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run",
            System.Reflection.BindingFlags.Default |
            System.Reflection.BindingFlags.InvokeMethod,
            null, oApp, oRunArgs);
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
