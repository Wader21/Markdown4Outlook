
using System;
using System.Drawing;
using System.ComponentModel;

namespace Markdown4Outlook
{
	/// <summary>
	/// Description of Class1.
	/// </summary>
	public class Configuration
	{
		public bool enableAddin { get; set; }
		
		public String editorFont { get; set; }
		
		public String styleFilePath { get; set; }
		
		public Configuration()
		{
			enableAddin = true;
			editorFont = null;
			styleFilePath = null;
		}
		
		public void setFont(Font font) {
			TypeConverter converter = TypeDescriptor.GetConverter(typeof(Font));
			editorFont = converter.ConvertToString(font);
		}
		
		public Font getFont() {
			if (editorFont != null) {
				TypeConverter converter = TypeDescriptor.GetConverter(typeof(Font));
				return (Font) converter.ConvertFromString(editorFont);
			} else {
				return null;
			}
		}
	}
}
