using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample
{
    public class SampleAddIn : IEplAddIn
    {
        public bool OnExit()
        {
            return true;
        }

        public bool OnInit()
        {
            return true;
        }

        public bool OnInitGui()
        {
            Eplan.EplApi.Gui.Menu sampleMenu = new Eplan.EplApi.Gui.Menu();
            Eplan.EplApi.Gui.Menu sampleMenu_Ch = new Eplan.EplApi.Gui.Menu();

            uint menuId = sampleMenu.AddMainMenu("[API Test Menu]", Eplan.EplApi.Gui.Menu.MainMenuName.eMainMenuHelp,
                                                  "Api Ext Sample", "ActionApiExtSimple", "Api Ext Sample", 1);
            
            menuId = sampleMenu.AddMenuItem("Call Other Action", "ActionApiExtCallOtherAction", "Call Other Action", menuId, 1, false, false);
            menuId = sampleMenu.AddMenuItem("Api Samples", "ActionApiExtSamples", "Api Samples", menuId, 1, false, false);
            menuId = sampleMenu.AddPopupMenuItem("PopupMenuName", "ChildMunuName", "ActionApiExtTest123", "PopupMenuName", menuId, 1, false, false);
      
            return true;
        }

   

        public bool OnRegister(ref bool bLoadOnStart)
        {
            return true;
        }

        public bool OnUnregister()
        {
            return true;
        }
    }
}
