using System;
using Microsoft.Office.Core;

namespace LMStudioAddin
{
    public partial class ThisAddIn
    {
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }
    }
}
