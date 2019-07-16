using Eplan.EplApi.ApplicationFramework;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample.Actions
{

    public class ActionTest : IEplAction
    {
        public bool Execute(ActionCallingContext ctx)
        {
            MessageBox.Show("ActionTest was called!");

            return true;
        }

        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "ActionTest";
            Ordinal = 20;

            return true;
        }

        /// <summary>
        /// Documentation function for the Action; is called of the system as required 
        /// Bescheibungstext delivers for the Action itself and if the Action String-parameters ("Kommandozeile")
        /// also name and description of the single parameters evaluates
        /// </summary>
        /// <param name="actionProperties"> This object must be filled with the information of the Action.</param>
        public void GetActionProperties(ref ActionProperties actionProperties)
        {
            // Description 1st parameter
            // ActionParameterProperties firstParam= new ActionParameterProperties();
            // firstParam.set("Param1", "1. Parameter for ActionApiExtSimple"); 
            // actionProperties.addParameter(firstParam);

            // Description 2nd parameter
            // ActionParameterProperties firstParam= new ActionParameterProperties();
            // firstParam.set("Param2", "2. Parameter for ActionApiExtSimple"); 
            // actionProperties.addParameter(firstParam);

        }
    }

}