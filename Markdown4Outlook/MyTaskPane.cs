using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Markdig;

namespace Markdown4Outlook
{
    public partial class MyTaskPane : UserControl , Outlook.Tools.ITaskPane
    {
    	private String styleFilePath;
    	
		#region Ctor
        
		public MyTaskPane()
        {
            InitializeComponent();
        }

		#endregion

		#region Properties
		
		private Addin ParentAddin { get; set; }

		#endregion
		
        #region ITaskpane

        public void OnConnection(Outlook.Application application, Office._CustomTaskPane definition, object[] customArguments)
        {
			if(customArguments.Length > 0)
				ParentAddin = customArguments[0] as Addin;
        }

        public void OnDisconnection()
        {

        }

        public void OnDockPositionChanged(MsoCTPDockPosition position)
        {
            
        }

        public void OnVisibleStateChanged(bool visible)
        {
			if(null != ParentAddin && null != ParentAddin.RibbonUI)
				ParentAddin.RibbonUI.InvalidateControl("tooglePaneVisibleButton");
        }
        
        public void updateConfig(Configuration config) {
        	Font font = config.getFont();
        	if (font != null) {
        		this.inputBox.Font = font;	
        	}
        	
        	this.styleFilePath = config.styleFilePath;
        }
        
		void UpateHTMLPreview(object sender, EventArgs e) {
        	
			var mailText = inputBox.Text;
				
			var markdownHTML = Markdown.ToHtml(mailText);
			
			var preMailer = new PreMailer.Net.PreMailer(markdownHTML);
			
			if (File.Exists(styleFilePath)) {
				
				var cssSource = File.ReadAllText(styleFilePath);

				var result = preMailer.MoveCssInline
	                (
	                    removeStyleElements: true,
	                    ignoreElements: null,
	                    css: cssSource,
	                    stripIdAndClassAttributes: true,
	                    removeComments: true
	                );			
				
				markdownHTML = result.Html;
			}
			
			previewBrowser.DocumentText = markdownHTML;		
		}
        

        #endregion
    }
}
