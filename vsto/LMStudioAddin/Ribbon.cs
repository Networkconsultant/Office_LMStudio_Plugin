using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace LMStudioAddin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public Ribbon() { }

        public string GetCustomUI(string ribbonID) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'
          onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab idMso='TabHome'>
        <group id='LMStudioGroup' label='LMStudio AI'>
          <button id='ToggleTaskPane'
                  label='Open AI Assistant'
                  imageMso='HappyFace'
                  size='large'
                  onAction='ToggleTaskPane_Click'
                  screentip='Open the LMStudio AI task pane'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void ToggleTaskPane_Click(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ToggleTaskPane();
        }
    }
}
