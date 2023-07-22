using System;

namespace DataJuggler.Excelerate.Interfaces
{
    public interface IExcelerateObject
    {

        #region Methods

            #region Load(Row row)
            /// <summary>
            /// This method loads a Address object from a Row.
            /// </Summary>
            /// <param name="row">The row which the row.Columns[x].ColumnValue will be used to load this object.</param>
            void Load(Row row);
            #endregion

            #region Save(Row row)
            /// <summary>
            /// This method saves a Address object back to a Row.
            /// </Summary>
            /// <param name="row">The row which the row.Columns[x].ColumnValue will be set to Save back to Excel.</param>
            Row Save(Row row);
            #endregion

        #endregion

        #region Properties

            #region ChangedColumns { get; set; }
            /// <summary>
            /// This string contains the column indexes that have changed since loading.
            /// This is used so only fields that have changes are updated
            /// </summary>
            string ChangedColumns { get; set; }
            #endregion

            #region Loading {get; set; }
            /// <summary>
            /// Is this object currently loading
            /// </summary>
            bool Loading {get; set; }
            #endregion

            #region RowId { get; set; }
            /// <summary>
            /// This unique identifier is used to find this row.
            /// </summary>
            Guid RowId { get; set; }
            #endregion

        #endregion

    }
}
