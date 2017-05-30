rm -r -f *.wixobj
rm -r -f *.wixpdb
rm -r -f *.msi

"C:\Program Files (x86)\WiX Toolset v3.11\bin\candle" setup.wxs
"C:\Program Files (x86)\WiX Toolset v3.11\bin\light" -cultures:en-US -ext WixUIExtension  -ext WixNetFxExtension -sice:ICE91 -sice:ICE57 -o Markdown4Outlook.msi setup.wixobj

@pause