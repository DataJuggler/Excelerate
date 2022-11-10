

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#endregion

namespace DataJuggler.Excelerate.Enumerations
{
    
    #region EditorTypeEnum : int
    /// <summary>
    /// This is used to set the editor type for the Grid when it goes into EditMode.
    /// </summary>
    public enum EditorTypeEnum : int
    {
        ReadOnly = 0,
        Text = 1,
        CheckBox = 2,
        ComboBox = 3
    }
    #endregion

}
