using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Configuration;
using System.Reflection;
using System.Text;
using System.IO;
using System.Web.Script.Serialization;

using NetOffice;
using NetOffice.Tools;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using NetOffice.OutlookApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;


using log4net;

namespace Markdown4Outlook
{
	[COMAddin("Markdown4Outlook", "Assembly Description", 3), ProgId("Markdown4Outlook.Addin"), Guid("B49C8937-2944-44BF-95A1-B08EA6ACE754")]
	[
		RegistryLocation(RegistrySaveLocation.CurrentUser)
		,CustomUI("Markdown4Outlook.RibbonUI.xml")
	]
	public class Addin : Outlook.Tools.COMAddin
	{
		private static readonly log4net.ILog log = log4net.LogManager.GetLogger
	    	(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		
		private Dictionary<NetOffice.OutlookApi.Inspector, InspectorWrapper>
			inspectorWrappers = new Dictionary<NetOffice.OutlookApi.Inspector, InspectorWrapper>();
		
		private Configuration config = new Configuration();
		
		private String configFilePath;
		
		public Addin()
		{
			this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
			this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);
		}

		internal Office.IRibbonUI RibbonUI { get; private set; }

		private void Addin_OnStartupComplete(ref Array custom)
		{
			initLog();
			
			loadConfig();
			
			var inspectors = Application.Inspectors as NetOffice.OutlookApi.Inspectors;
			inspectors.NewInspectorEvent += Inspectors_NewInspectorEvent;
			
			Application.QuitEvent += new Outlook.Application_QuitEventHandler(Application_OnQuit);
		}
		
		private void Application_OnQuit()
		{
			saveConfig();
		}

		private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
		}
		
		public object GetImage(string imageName) {
			
			var sb = new StringBuilder();
	        sb.Append(Assembly.GetExecutingAssembly().GetName().Name);
	        sb.Append(".Resources.");
	        sb.Append(imageName);
	        Stream imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(sb.ToString());
			
	        return new Bitmap(imageStream);
		}
		
        public void About_Click(Office.IRibbonControl control)
        {
			MessageBox.Show(
        		String.Format("{0} Version {1}",
        		              Constants.ADDIN_TITLE, this.GetType().Assembly.GetName().Version),
				Constants.ADDIN_TITLE, 
				MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /* callback function to determine the toggle state of the button */
		public bool ToogleAddInEnableButton_GetPressed(Office.IRibbonControl control)
		{
			return config.enableAddin;
		}

		public void ToogleAddInEnableButton_Click(Office.IRibbonControl control, bool pressed)
		{
			config.enableAddin = pressed;
			
			foreach (var item in inspectorWrappers) {
				var inspectorWrapper = item.Value;
				inspectorWrapper.TaskPane.Visible = pressed;
        	}
		}

		public void OnLoadRibonUI(Office.IRibbonUI ribbonUI)
        {
			RibbonUI = ribbonUI;
        }

		protected override void OnError(ErrorMethodKind methodKind, System.Exception exception)
		{
			MessageBox.Show("An error occurend in " + methodKind.ToString(), "Markdown4Outlook");
		}

		[RegisterErrorHandler]
		public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, System.Exception exception)
		{
			MessageBox.Show("An error occurend in " + methodKind.ToString(), "Markdown4Outlook");
		}
		
		public override string GetCustomUI(string RibbonID)
        {
			if (RibbonID != Constants.TARGET_RIBBON_ID) {
				return "";
			} 
				
            var ui = base.GetCustomUI(RibbonID);
            return ui;
        }
		
		private String getConfigFilePath(){
			return System.IO.Path.Combine(
				Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
			    Constants.CONFIG_FILE_NAME);
		}
		
		private void loadConfig() {
			configFilePath = getConfigFilePath();
			
			log.Info("Load config from :" + configFilePath);
			
			if (File.Exists(configFilePath)) {
				config = (new JavaScriptSerializer()).Deserialize<Configuration>(File.ReadAllText(configFilePath));				
				log.Info("config :" + config);
			}
		}
		
		private void saveConfig() {
			log.Info("Save config to :" + configFilePath);
			
			File.WriteAllText(configFilePath, (new JavaScriptSerializer()).Serialize(config));
		}
		
		private void initLog() {
			var sb = new StringBuilder();
	        sb.Append(Assembly.GetExecutingAssembly().GetName().Name);
	        sb.Append(".");
	        sb.Append("log4net.config");
	        Stream configStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(sb.ToString());
	        log4net.Config.XmlConfigurator.Configure(configStream);
	        
			log.Info(String.Format("Addin started in Outlook Version {0}", Application.Version));
		}
		
		private void Inspectors_NewInspectorEvent(NetOffice.OutlookApi._Inspector _inspector)
	    {
			
			var inspector = _inspector as NetOffice.OutlookApi.Inspector;
			
	        var ai = inspector.CurrentItem as Outlook.MailItem;
	        if (ai == null) {
	        	return;
	        }

			log.Info("Create new taskpane with config:" + config);
	        
	        var taskPane = TaskPaneFactory.CreateCTP(typeof(MyTaskPane).FullName, Constants.TASK_PANE_TITLE, inspector) as Office.CustomTaskPane;
	        taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
	        taskPane.Width = 500;
	        
	        taskPane.Visible = config.enableAddin;
	        
	        var myTaskPane = (MyTaskPane)taskPane.ContentControl;
	        
			log.Info("myTaskPane is" + myTaskPane);
	        
			myTaskPane.updateConfig(config);
	        
	        
	        var wrapper = new InspectorWrapper(inspector,taskPane, this);
	        inspectorWrappers.Add(inspector, wrapper);
	    }
		
		public void inspectorClosed(InspectorWrapper inspectorWrapper) {
        	inspectorWrappers.Remove(inspectorWrapper.Inspector);
		}
		
		public void ConfigFont_Click(Office.IRibbonControl control)
		{
			var configFontDialog = new FontDialog();
			
			if (config.editorFont != null) {
				configFontDialog.Font = config.getFont();
			}
			
			if (configFontDialog.ShowDialog() == DialogResult.OK) {
				config.setFont(configFontDialog.Font);
				updateTaskPaneConfig();
			}
		}

		public void ConfigStyle_Click(Office.IRibbonControl control)
		{
			var configStyleDialog = new OpenFileDialog();
			configStyleDialog.Multiselect = false;
			
			var styleFilePath = config.styleFilePath;
			if (styleFilePath != null && File.Exists(styleFilePath)) {
				configStyleDialog.InitialDirectory = System.IO.Path.GetDirectoryName(styleFilePath);
				configStyleDialog.FileName = System.IO.Path.GetFileName(styleFilePath);
			}
			
			if (configStyleDialog.ShowDialog() == DialogResult.OK) {
				config.styleFilePath = configStyleDialog.FileName; 
				updateTaskPaneConfig();
			}
		}
		
		private void updateTaskPaneConfig() {
			foreach (var item in inspectorWrappers) {
				var inspectorWrapper = item.Value;
				var taskPane =inspectorWrapper.TaskPane;
				var myTaskPane = (MyTaskPane)taskPane.ContentControl;
				myTaskPane.updateConfig(config);
	    	}
		}
		
		public void Feedback_Click(Office.IRibbonControl control)
		{
            var mailItem = (Outlook.MailItem)
                this.Application.CreateItem(NetOffice.OutlookApi.Enums.OlItemType.olMailItem);
            mailItem.Subject = Constants.FEEDBACK_MAIL_SUBJECT;
            mailItem.To = Constants.FEEDBACK_MAIL_TO;
            mailItem.Body = Constants.FEEDBACK_MAIL_BODY;
            mailItem.Display(false);			
		}
		
		
    }
}

