
using System;
using Outlook = NetOffice.OutlookApi;
using Office = NetOffice.OfficeApi;

namespace Markdown4Outlook
{
	/// <summary>
	/// Description of Class1.
	/// </summary>
	public class InspectorWrapper
	{

		private NetOffice.OutlookApi.Inspector inspector;
		private Office.CustomTaskPane taskPane;
		private Addin addIn;

		public InspectorWrapper(NetOffice.OutlookApi.Inspector inspector, Office.CustomTaskPane taskPane, Addin addIn)
		{
			this.inspector = inspector;
			this.taskPane = taskPane;
			this.addIn = addIn;

			inspector.CloseEvent += Inspector_CloseEvent;
		}

		private void Inspector_CloseEvent() {
			if (taskPane != null) {
				taskPane.Delete();
 			}
 			taskPane = null;

 			addIn.inspectorClosed(this);

			inspector.CloseEvent -= Inspector_CloseEvent;
 			inspector = null;
 		}

		public  NetOffice.OutlookApi.Inspector Inspector {
			get {
				return inspector;
			}
		 }

		public Office.CustomTaskPane TaskPane {
			get {
				return taskPane;
			}
		 }

	}
}
