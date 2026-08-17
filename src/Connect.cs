using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

[assembly: ComVisible(true)]
[assembly: Guid("9478E440-4F1D-4A58-AC28-C01C7F88B9E5")]
[assembly: AssemblyTitle("Word Move to Heading")]
[assembly: AssemblyVersion("2.2.0.0")]
[assembly: AssemblyFileVersion("2.2.0.0")]

namespace WordMoveToHeading
{
    [ComVisible(true)]
    [Guid("AD9CF34E-04AB-4628-B2E8-90CA487BC348")]
    [ProgId("WordMoveToHeading.Connect2")]
    // Ribbon callbacks are resolved by name through IDispatch. AutoDispatch
    // exposes the public callback methods without exposing fields.
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private MoveToHeadingMenu _menu;
        private Office.IRibbonUI _ribbon;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            var word = application as Word.Application;
            if (word != null && _menu == null)
            {
                _menu = new MoveToHeadingMenu(word);
                Office.COMAddIn comAddIn = addInInst as Office.COMAddIn;
                if (comAddIn != null)
                {
                    comAddIn.Object = this;
                }
            }
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""WordMoveToHeading.OutlineTools"" label=""Outline tools"">
          <button id=""WordMoveToHeading.DetectHeadings""
                  label=""Nhận diện Heading""
                  size=""large""
                  screentip=""Tự động gán Outline level""
                  supertip=""Nhận diện I., II., 1., 1.1., a)... và gán Outline level tương ứng.""
                  onAction=""OnDetectHeadings"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuText"">
      <dynamicMenu id=""WordMoveToHeading.DynamicMenu""
                   label=""Move to""
                   getContent=""GetMoveToMenuContent"" />
    </contextMenu>
  </contextMenus>
</customUI>";
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public string GetMoveToMenuContent(Office.IRibbonControl control)
        {
            return _menu == null
                ? "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><button id=\"WordMoveToHeading.NotReady\" label=\"Add-in is starting...\" enabled=\"false\" /></menu>"
                : _menu.BuildDynamicMenuXml();
        }

        // Used by the installer smoke test; Ribbon uses GetMoveToMenuContent.
        public string DiagnosticGetMoveToMenuContent()
        {
            return GetMoveToMenuContent(null);
        }

        public void OnMoveToHeading(Office.IRibbonControl control)
        {
            int headingStart;
            if (_menu != null && control != null && int.TryParse(control.Tag, out headingStart))
            {
                _menu.MoveSelectionToHeading(headingStart);
                if (_ribbon != null)
                {
                    _ribbon.InvalidateControl("WordMoveToHeading.DynamicMenu");
                }
            }
        }

        public void OnDetectHeadings(Office.IRibbonControl control)
        {
            if (_menu != null)
            {
                _menu.DetectAndAssignOutlineLevels();
                if (_ribbon != null)
                {
                    _ribbon.InvalidateControl("WordMoveToHeading.DynamicMenu");
                }
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            DisposeMenu();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            DisposeMenu();
        }

        private void DisposeMenu()
        {
            if (_menu != null)
            {
                _menu.Dispose();
                _menu = null;
            }
            _ribbon = null;
        }
    }
}

