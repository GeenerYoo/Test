
using System;
using Eplan.EplApi.ApplicationFramework;

namespace Addin_Test
{
    /// <summary>
    ///   That is an example for a EPLAN addin.  
    ///   Exactly a class must implement the interface Eplan.EplApi.ApplicationFramework.IEplAddIn.  
    ///   An Assembly is identified through this criterion as EPLAN addin!  
    /// </summary>  
    public class AddInModule : IEplAddIn
    {
        /// <summary>
        /// The function is called once during registration add-in.
        /// </summary>
        /// <param name="bLoadOnStart"> true: In the next Eplan session, add-in will be loaded during initialization</param>
        /// <returns></returns>
        public bool OnRegister(ref System.Boolean bLoadOnStart)
        {
            bLoadOnStart = true;

            return true;
        }
        /// <summary>
        /// The function is called during unregistration the add-in.
        /// </summary>
        /// <returns></returns>
        public bool OnUnregister()
        {
            return true;
        }

        /// <summary>
        /// The function is called during Eplan initialization or registration the add-in.  
        /// </summary>
        /// <returns></returns>
        public bool OnInit()
        {

            return true;

        }
        /// <summary>
        /// The function is called during Eplan initialization or registration the add-in, when GUI was already initialized and add-in can modify it. 
        /// </summary>
        /// <returns></returns>
        public bool OnInitGui()
        {
            Eplan.EplApi.Gui.Menu sampleMenu = new Eplan.EplApi.Gui.Menu();
            Eplan.EplApi.Gui.Menu sampleMenu_Child = new Eplan.EplApi.Gui.Menu();

            uint menuId = sampleMenu.AddMainMenu("[EPLAN API]", Eplan.EplApi.Gui.Menu.MainMenuName.eMainMenuHelp,"FirstMenu", "ActionTest", "1st", 1);

            uint  menuId2 = sampleMenu.AddPopupMenuItem("PopupMenuName", "ChildMenuName", "ActionPopupChildMenu1", "PopupMenuName", menuId, 1, false, false);

           

            return true;

        }
        /// <summary>
        /// This function is called during closing Eplan or unregistration the add-in. 
        /// </summary>
        /// <returns></returns>
        public bool OnExit()
        {
            return true;
        }

    }
}
