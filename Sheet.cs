using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using ExcelDataReader;                                  // For doing things with a thing
using VC = MastersHelperLibrary.VirtualConsole;



namespace FTR_Utils
{
    public class Sheet
    {
        public DataSet result;
        public DataTable table;
        public bool bFileLoaded;
        public int numOfTables;
        public int numOfRows;
        public int numOfCols;

        public Dictionary<string, int> columnNames = new Dictionary<string, int>();

        public string DefaultFileName;
        public string TestFileName = "testFile.xlsx";
        public Sheet()
        {
            DefaultFileName = TestFileName;
            InitVC();
        }
        public Sheet(string s)
        {
            DefaultFileName = s;
            InitVC();
        }

        public void InitVC()
        {
            VC.AddNewConsoleCommand(VC_LoadTable, "LoadTable", "LoadTable /myFileName.xlsx  - load a .xlsx file to a DataTable.");
            VC.AddNewConsoleCommand(VC_PrintTable, "PrintTable", "Prints the freaking table");
            VC.AddNewConsoleCommand(VC_PrintRow, "PrintRow", "Prints a freaking row");
            VC.AddNewConsoleCommand(VC_PrintColumn, "PrintColumn", "Prints a freaking column");
            VC.AddNewConsoleCommand(VC_PrintNode, "PrintNode", "Prints a freaking node");
            VC.AddNewConsoleCommand(VC_SetNode, "SetNode", "Sets a freaking node");
            VC.AddNewConsoleCommand(VC_PrintDims, "PrintDims", "Return the table dimensions as ROWSxCOLS");
            VC.AddNewConsoleCommand(VC_PrintColumnNames, "PrintColumnNames", "Return the table of column names, and their numerical order");
        }

        public string GetErr()
        {
            return GetErr(0);
        }
        public string GetErr(int i)
        {
            string s = "";
            switch (i)
            {
                case 0:
                    s = "*** generic and frustratingly nonspecific error message ***";
                    break;
                case 1:
                    s = "*** \x0d\x0a Directory is not initialized. *** \x0d\x0a Try 'LoadTable [filename]'.";
                    break;
                case 2:
                    s = "*** index out of range ***";
                    break;
                case 3:
                    s = "*** the arg needs to be an integer *** \x0d\x0a e.g. 'PrintRow 5'";
                    break;
                case 4:
                    s = "*** both args need to be integers *** \x0d\x0a e.g. 'PrintNode 5 2'";
                    break;
                case 5:
                    s = "*** PrintNode takes exactly 2 integer args ***";
                    break;
                case 6:
                    s = "*** columnNames key not found ***";
                    break;
                default:
                    s = "*** GetErr switch default - What a trivial usage of your time. ***";
                    break;
            }
            return s;
        }

        public string LoadFile()
        {
            return LoadFile(DefaultFileName);
        }
        public string LoadFile(string s)
        {
            try
            {
                using (var stream = File.Open(s, FileMode.Open, FileAccess.Read))
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                        result = reader.AsDataSet();

                numOfTables = result.Tables.Count;
                if (numOfTables > 0)
                    table = result.Tables[0];
                    numOfRows = table.Rows.Count;
                    if (table.Columns.Count > numOfCols)
                        numOfCols = table.Columns.Count;
                    bFileLoaded = true;

                return PrintTableData();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        public string SetTableNode(int i, int j, string s)
        {
            if (!bFileLoaded)
                return "";
            try
            {
                if (i < numOfRows && j < numOfCols)
                {
                    table.Rows[i].ItemArray[j] = s;
                    return String.Format("*** set row.{0} col.{1} = {2}", i, j, s);
                }
                else
                    return String.Format("*** out of bounds index ***\x0d\x0a numOfRows={0}, numOfCols={1}\x0d\x0a (zero-based)");
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }


        ///////////////////////////////////////////////////////////////////
        ///Get / Print
        /////////////////////////////////////////////////////////////////////

        public string PrintColumnNames()
        {
            StringBuilder output = new StringBuilder();
            output.AppendFormat("Number of columns found: {0}\x0d\x0a", numOfCols);
            foreach (KeyValuePair<string, int> kv in columnNames)
                output.AppendFormat("{0} : {1}\x0d\x0a", kv.Key.PadLeft(20), kv.Value);
            output.Append("\x0d\x0a");
            return output.ToString();
        }

        public string PrintTableDims()
        {
            if (!bFileLoaded)
                return GetErr(1);

            return String.Format("Table dimensions: {0}x{1} \x0d\x0a Reference by zero-based indexing", numOfRows, numOfCols);
        }
        public string PrintTableData()
        {
            if (!bFileLoaded)
                return GetErr(1);

            StringBuilder output = new StringBuilder();

            output.AppendFormat("Entries found: {0}\x0d\x0a\x0d\x0a", numOfRows);
            foreach (DataRow r in table.Rows)
            {
                output.Append(PrintTableRow(r));
                output.Append("\x0d\x0a");
            }
            return output.ToString();
        }

        public string PrintTableRow(int i)
        {
            if (!bFileLoaded)
                return GetErr(1);
            if (i < numOfRows)
                return PrintTableRow(table.Rows[i]);
            else
                return GetErr(2);
        }
        public string PrintTableRow(DataRow r)
        {
            if (!bFileLoaded)
                return GetErr(1);

            StringBuilder output = new StringBuilder();
            foreach (var c in r.ItemArray)
                output.AppendFormat("{0} ", c);
            //output.Append("\x0d\x0a");
            return output.ToString();
        }

        public string PrintTableColumn(string s)
        {
            if (!bFileLoaded)
                return GetErr(1);
            if (!columnNames.TryGetValue(s.Trim().ToLower(), out int i))
                return GetErr(6);
            return PrintTableColumn(columnNames[s.Trim().ToLower()]);
        }
        public string PrintTableColumn(int i)
        {
            if (!bFileLoaded)
                return GetErr(1);

            StringBuilder output = new StringBuilder();
            output.Append("\x0d\x0a");
            foreach (DataRow r in table.Rows)
                output.AppendFormat("{0}\x0d\x0a", r.ItemArray[i].ToString().PadLeft(50));
            output.Append("\x0d\x0a");

            return output.ToString();
        }

        public string PrintTableNode(int i, int j)
        {
            if (!bFileLoaded)
                return GetErr(1);
            if (i >= numOfRows || j >= numOfCols)
                return GetErr(2);
            return table.Rows[i].ItemArray[j].ToString();
        }



        ////////////////////////////////////////////////////////////////////////
        //VC Commands
        /// ////////////////////////////////////////////////////////////////////

        public string VC_LoadTable(string s)
        {
            if (s.Trim().Length == 0)
                VC.Send(LoadFile());
            else
                VC.Send(LoadFile(s.Trim()));
            
            return "";
        }
        public string VC_PrintDims(string s)
        {
            if (!bFileLoaded)
            {
                VC.Send(GetErr(1));
                return "";
            }

            VC.Send(PrintTableDims());

            return "";
        }
        public string VC_PrintColumnNames(string s)
        {
            if (!bFileLoaded)
            {
                VC.Send(GetErr(1));
                return "";
            }

            VC.Send(PrintColumnNames());
            return "";
        }
        public string VC_PrintTable(string s)
        {
            if (!bFileLoaded)
            {
                VC.Send(GetErr(1));
                return "";
            }
            VC.Send(PrintTableData());
            return "";
        }
        public string VC_PrintRow(string s)
        {
            if (!bFileLoaded)
            {
                VC.Send(GetErr(1));
                return "";
            }
            if (int.TryParse(s.Trim(), out int i))
                VC.Send(PrintTableRow(i));
            else
                VC.Send(GetErr(3));

            return "";
        }
        public string VC_PrintColumn(string s)
        {
            if (!bFileLoaded)
            {
                VC.Send(GetErr(1));
                return "";
            }

            if (int.TryParse(s.Trim(), out int i))
                VC.Send(PrintTableColumn(i));
            else if (columnNames.ContainsKey(s.Trim().ToLower()))
                VC.Send(PrintTableColumn(columnNames[s.Trim().ToLower()]));
            else
                VC.Send(GetErr(3));

            return "";
        }
        public string VC_PrintNode(string s)
        {
            if (!bFileLoaded)
            {
                VC.Send(GetErr(1));
                return "";
            }
            string[] sArgs = s.Trim().Split();
            if (sArgs.Length == 2)
                if (int.TryParse(sArgs[0].Trim(), out int i) && int.TryParse(sArgs[1].Trim(), out int j))
                    VC.Send(PrintTableNode(i, j));
                else
                    VC.Send(GetErr(4));
            else
                VC.Send(GetErr(5));
            return "";
        }
        public string VC_SetNode(string s)
        {
            if (!bFileLoaded)
            {
                VC.Send(GetErr(1));
                return "";
            }

            string[] sArgs = s.Trim().Split();
            if (sArgs.Length == 3)
            {
                if (int.TryParse(sArgs[0], out int i) && int.TryParse(sArgs[1], out int j))
                {
                    VC.Send(SetTableNode(i, j, s));
                }
            }
            else
                VC.Send(GetErr(6));
            return "";
        }

    }
}