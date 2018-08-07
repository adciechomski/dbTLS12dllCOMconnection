using dbTLS12dllCOMconnection;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System;

[ComVisible(true)]
public interface IAddInUtilities
{
    string[,] ImportData(string storedProc, string SqlConString, string procParameter);
    string[,] GetArrayFromCSharp();
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public class AddInUtilities : IAddInUtilities
{
    // private string storedProc = "";
    // private string SqlConString = "";
    public string[,] ImportData(string storedProc, string SqlConString, string procParameter)
    {
        Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
        SqlConnection sqlConnection1 = new SqlConnection(SqlConString);
        SqlDataAdapter sda = new SqlDataAdapter(storedProc, sqlConnection1);
        sda.SelectCommand.CommandType = CommandType.StoredProcedure;
        sda.SelectCommand.Parameters.Add(new SqlParameter("@pStrStatus", SqlDbType.VarChar));
        sda.SelectCommand.Parameters["@pStrStatus"].Value = procParameter;
        DataTable dt = new DataTable();
        sda.Fill(dt);
        sqlConnection1.Close();
        return ConvertDataTableToArray(dt);
    }

    public string[,] ConvertDataTableToArray(DataTable dt)
    {
        //Declare 2D-String array
        string[,] arrString = new string[dt.Rows.Count + 1, dt.Columns.Count];
        int index = 0;

        //Save ColumnName in 1st row of the array
        foreach (DataColumn dc in dt.Columns)
        {
            arrString[0, index] = Convert.ToString(dc.ColumnName);
            index++;
        }

        index = 0;
        //DataTable values in array,
        for (int row = 1; row < dt.Rows.Count + 1; row++)
        {
            for (int col = 0; col < dt.Columns.Count; col++)
            {
                int roww = row - 1;
                arrString[row, col] = Convert.ToString(dt.Rows[roww][col]);
            }
        }
        return arrString;
    }

    //test method to be deleted
    public string[,] GetArrayFromCSharp()
    {
        string[,] array = new string[,]
{
            {"cat", "dog"},
            {"bird", "fish"},
};
        return array;
    }

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {


    }
}
