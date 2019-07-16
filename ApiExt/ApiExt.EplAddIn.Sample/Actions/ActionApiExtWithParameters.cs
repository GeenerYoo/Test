using Eplan.EplApi.ApplicationFramework;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample.Actions
{

    /// <summary>
    /// This class implements a EPLAN action.  The Action will register the Addins in that  <seealso cref="IEplAddIn.OnRegister"/> Registerst.
    /// <seealso cref="Eplan::EplApi::ApplicationFramework::IEplAction"/> 
    /// </summary>
    public class ActionApiExtWithParameters : IEplAction
    {
        /// <summary>
        /// Execution of the Action.  
        /// </summary>
        /// <returns>True:  Execution of the Action was successful</returns>
        public bool Execute(ActionCallingContext ctx)
        {
            string param1 = null;
            ctx.GetParameter("Param1", ref param1);

            string param2 = null;
            ctx.GetParameter("Param2", ref param2);

            string param3 = null;
            ctx.GetParameter("Param3", ref param3);

            MessageBox.Show(string.Format("ActionApiExtWithParameters\n[Param1: {0}], [Param2: {1}], [Param3: {2}]\nwas called!", param1, param2, param3));

            ctx.AddParameter("ReturnParam", "Action ReturnParam!!!");

            return true;
        }

        /// <summary>
        /// Function is called through the ApplicationFramework
        /// </summary>
        /// <param name="Name">Under this name, this Action in the system is registered</param>
        /// <param name="Ordinal">With this overload priority, this Action is registered</param>
        /// <returns>true: the return parameters are valid</returns>
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "ActionApiExtWithParameters";
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
            // firstParam.set("Param1", "1. Parameter for ActionApiExtWithParameters"); 
            // actionProperties.addParameter(firstParam);

            // Description 2nd parameter
            // ActionParameterProperties firstParam= new ActionParameterProperties();
            // firstParam.set("Param2", "2. Parameter for ActionApiExtWithParameters"); 
            // actionProperties.addParameter(firstParam);

        }

    }

}