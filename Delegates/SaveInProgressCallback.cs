
#region using statements

#endregion

namespace DataJuggler.Excelerate.Delegates
{

    
    #region delegate void SaveInProgressCallback(SaveWorksheetResponse response);
    /// <summary>
    /// This delegate is used to send information back to a client about a SaveaWorksheet
    /// operation.
    /// </summary>
    /// <param name="response"></param>
    public delegate void SaveInProgressCallback(SaveWorksheetResponse response);
    #endregion

}
