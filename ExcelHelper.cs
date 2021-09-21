

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

#endregion

namespace DataJuggler.Excelerate
{

    #region class ExcelHelper
    /// <summary>
    /// method [Enter Method Description]
    /// </summary>
    public class ExcelHelper
    {

        #region Methods

            #region GetColumnLetter(int column)
            /// <summary>
            /// returns the Column Letter for the column index (1 = A, 2 = B, 27 = AA, 78 = "ZZZ" I think)
            /// </summary>
            public static string GetColumnLetter(int column)
            {
                // initial value
                string columnLetter = ExcelCellAddress.GetColumnLetter(column);
                
                // return value
                return columnLetter;
            }
            #endregion
            
        #endregion

    }
    #endregion

}
