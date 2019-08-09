//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.Data.OleDb;
//using System.Data.OracleClient;
//using System.Data.SqlClient;
//using System.Text;
//using FirebirdSql.Data.FirebirdClient;
//using SKF.RS.STB.Generic;
////using Oracle.DataAccess.Client;

//namespace SKF.RS.STB.DB
//{
//    public class TableColumn
//    {
//        public string Name = null;
//        public object Value = null;

//        public TableColumn() { }
//        public TableColumn(string Name, object Value) { this.Name = Name; this.Value = Value; }
//    }
    
//    public enum DBType
//    {
//        None,
//        Oracle = 1,
//        MSSQL = 2,
//        Access = 3,
//        Firebird = 4
//    }

//    public class DBError
//    {
//        private bool _Initialized = false;
//        public bool Initialized { get { return _Initialized;}}

//        private DateTime _TimeStamp = DateTime.Now;
//        public DateTime TimeStamp { get { return (_Initialized ? _TimeStamp : new DateTime()); } }
        
//        private int _Code = 0;
//        public int Code { get { return (_Initialized ? _Code : 0); } }

//        private string _Message = string.Empty;
//        public string Message { get { return (_Initialized ? _Message : string.Empty); } }

//        public DBError() { }
//        public DBError(int Code, string Message)
//        {
//            _Code = Code;
//            _Message = Message;
//            GenericTools.DebugMsg(Message);
//            _Initialized = true;
//        }

//    }

//    public class DBTools
//    {
//        private OracleConnection OraConnection;
//        private SqlConnection SQLConnection = new SqlConnection();
//        private OleDbConnection MSAConnection = new OleDbConnection();
//        private FbConnection FBConnection = new FbConnection();

//        private bool _IsConnected = false;
//        public bool IsConnected { get { return _IsConnected; } }

//        public string Owner { get; set; }
//        public string ConnectionString { get; set; }
//        public string DBName { get; set; }
//        public string InitialCatalog { get; set; }
//        public string User = "SKFUser1";
//        public string Password = "cm";
//        public DBType DBType = DBType.None;
////        private bool Debug2 = Debug;
//        public static string DBVersion { get; set; }

//        public void xxx_SetConnected(bool Connected) { _IsConnected = Connected; }

//        public void DebugMsg(string DebugMessage) { GenericTools.DebugMsg(DebugMessage); }

//        public uint SQLtoUInt(string Column, string Table, string Where)
//        {
//            if (string.IsNullOrEmpty(Column) | string.IsNullOrEmpty(Table)) { return uint.MinValue; }
//            string Query = string.Empty;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    Query = "select abs(ceil(round(" + Column + ",0))) from " + Owner + Table;
//                    break;

//                case DBType.MSSQL:
//                    Query = "select abs(cast(round(" + Column + ",0) as int)) from " + Owner + Table;
//                    break;

//                case DBType.Access:
//                    Query = "select abs(int(round(" + Column + ",0))) from " + Owner + Table;
//                    break;

//                default:
//                    Query = "select " + Column + " from " + Owner + Table;
//                    break;
//            }
//            if (!string.IsNullOrEmpty(Where)) { Query += " where " + Where; }
//            return SQLtoUInt(Query);
//        }
//        public uint SQLtoUInt(string Column, string Table)
//        {
//            return SQLtoUInt(Column, Table, "");
//        }
//        public uint SQLtoUInt(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            uint ReturnValue = 0;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    ReturnValue = SQLtoUInt(OraConnection, sSQLQuery);
//                    break;

//                case DBType.MSSQL:
//                    ReturnValue = SQLtoUInt(SQLConnection, sSQLQuery);
//                    break;

//                case DBType.Access:
//                    ReturnValue = SQLtoUInt(MSAConnection, sSQLQuery);
//                    break;

//                case DBType.Firebird:
//                    ReturnValue = SQLtoUInt(FBConnection, sSQLQuery);
//                    break;
//            }
//            DebugMsg("SQLtoUInt(" + sSQLQuery + ") = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//        private uint SQLtoUInt(OracleConnection oSQLConnection, string sSQLQuery)
//        {
//            uint ReturnValue = uint.MinValue;
//            if (IsConnected)
//            {
//                /*
//                DataTable TableResult = DataTable(oSQLConnection, sSQLQuery);
//                if (TableResult.Rows.Count > 0) ReturnValue = Convert.ToInt32(TableResult.Rows[0][0]);
//                */

//                OracleCommand Command = new OracleCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Convert.ToUInt32(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    //if (ex.InnerException != null) 
//                    DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            else
//            {
//                DebugMsg("SQLtoUInt(" + sSQLQuery + ") error: not connected to DB");
//            }
//            return ReturnValue;
//        }
//        private uint SQLtoUInt(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            uint ReturnValue = uint.MinValue;
//            if (IsConnected)
//            {
//                SqlCommand Command = new SqlCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Convert.ToUInt32(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            else
//            {
//                DebugMsg("SQLtoUInt(" + sSQLQuery + ") error: not connected to DB");
//            }
//            return ReturnValue;
//        }
//        private uint SQLtoUInt(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            uint ReturnValue = uint.MinValue;
//            if (IsConnected)
//            {
//                OleDbCommand Command = new OleDbCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Convert.ToUInt32(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            else
//            {
//                DebugMsg("SQLtoUInt(" + sSQLQuery + ") error: not connected to DB");
//            }
//            return ReturnValue;
//        }

//        // ADICIONADO POR MATEUS
//        // DATA: 04/04/2014
//        // EXPANSÃO DA BIBLIOTECA PARA FIREBIRD
//        // PROJETO IMPLANTAÇÃO @DS NO CDR
//        private uint SQLtoUInt(FbConnection oSQLConnection, string sSQLQuery)
//        {
//            DataTable oRecordSet_DataTable = new DataTable();
//            try
//            {
//                FbDataAdapter oRecordSet_DataAdapter = new FbDataAdapter(sSQLQuery, oSQLConnection);
//                try
//                {
//                    oRecordSet_DataAdapter.Fill(oRecordSet_DataTable);
//                    oRecordSet_DataAdapter.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (oRecordSet_DataAdapter != null) oRecordSet_DataAdapter.Dispose();
//                    if (ex.InnerException != null) DebugMsg("RecordSet:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            catch (Exception ex)
//            {
//                DebugMsg("RecordSet:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//            }

//            return uint.Parse(oRecordSet_DataTable.Rows[0][0].ToString());
//        }

//        public Int32 SQLtoInt(string Field, string Table, string Where)
//        {
//            if (string.IsNullOrEmpty(Field) | string.IsNullOrEmpty(Table)) { return int.MinValue; }
//            string Query = string.Empty;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    Query = "select ceil(round(" + Field + ",0)) from " + Owner + Table;
//                    break;

//                case DBType.MSSQL:
//                    Query = "select cast(round(" + Field + ",0) as int) from " + Owner + Table;
//                    break;

//                case DBType.Firebird:
//                    Query = "select cast(" + Field + " as int) from " + Owner + Table;
//                    break;

//                case DBType.Access:
//                    Query = "select int(round(" + Field + ",0)) from " + Owner + Table;
//                    break;

//                default:
//                    Query = "select " + Field + " from " + Owner + Table;
//                    break;
//            }
//            if (!string.IsNullOrEmpty(Where)) { Query += " where " + Where; }
//            return SQLtoInt(Query);
//        }
//        public Int32 SQLtoInt(string Field, string Table)
//        {
//            return SQLtoInt(Field, Table, "");
//        }
//        public Int32 SQLtoInt(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            int ReturnValue = 0;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    ReturnValue = SQLtoInt(OraConnection, sSQLQuery);
//                    break;

//                case DBType.MSSQL:
//                    ReturnValue = SQLtoInt(SQLConnection, sSQLQuery);
//                    break;

//                case DBType.Access:
//                    ReturnValue = SQLtoInt(MSAConnection, sSQLQuery);
//                    break;

//                case DBType.Firebird:
//                    ReturnValue = SQLtoInt(FBConnection, sSQLQuery);
//                    break;
//            }
//            DebugMsg("SQLtoInt(" + sSQLQuery + ") = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }

//        /*
//        private Int32 SQLtoInt(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            return Convert.ToInt32(SQLtoFloat(oSQLConnection, sSQLQuery));
//        } 
//        */
//        private Int32 SQLtoInt(OracleConnection oSQLConnection, string sSQLQuery)
//        {
//            int ReturnValue = int.MinValue;
//            if (IsConnected)
//            {
//                /*
//                DataTable TableResult = DataTable(oSQLConnection, sSQLQuery);
//                if (TableResult.Rows.Count > 0) ReturnValue = Convert.ToInt32(TableResult.Rows[0][0]);
//                */

//                OracleCommand Command = new OracleCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Convert.ToInt32(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    //if (ex.InnerException != null) 
//                        DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            else
//            {
//                DebugMsg("SQLtoInt(" + sSQLQuery + ") error: not connected to DB");
//            }
//            return ReturnValue;
//        }
//        private Int32 SQLtoInt(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            int ReturnValue = int.MinValue;
//            if (IsConnected)
//            {
//                SqlCommand Command = new SqlCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Convert.ToInt32(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            else
//            {
//                DebugMsg("SQLtoInt(" + sSQLQuery + ") error: not connected to DB");
//            }
//            return ReturnValue;
//        }
//        private Int32 SQLtoInt(FbConnection oSQLConnection, string sSQLQuery)
//        {
//            int ReturnValue = int.MinValue;
//            if (IsConnected)
//            {
//                FbCommand Command = new FbCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Convert.ToInt32(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            else
//            {
//                DebugMsg("SQLtoInt(" + sSQLQuery + ") error: not connected to DB");
//            }
//            return ReturnValue;
//        }
//        private Int32 SQLtoInt(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            int ReturnValue = int.MinValue;
//            if (IsConnected)
//            {
//                OleDbCommand Command = new OleDbCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Convert.ToInt32(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            else
//            {
//                DebugMsg("SQLtoInt(" + sSQLQuery + ") error: not connected to DB");
//            }
//            return ReturnValue;
//        }

//        public float SQLtoFloat(string Field, string Table, string Where)
//        {
//            if (string.IsNullOrEmpty(Field) | string.IsNullOrEmpty(Table)) { return float.NaN; }
//            string Query = "select round(" + Field + ", 15) from " + Owner + Table;
//            if (!string.IsNullOrEmpty(Where)) { Query += " where " + Where; }
//            return SQLtoFloat(Query);
//        }
//        public float SQLtoFloat(string Field, string Table)
//        {
//            return SQLtoFloat(Field, Table, "");
//        }
//        public float SQLtoFloat(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            float ReturnValue = float.NaN;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    ReturnValue = SQLtoFloat(OraConnection, sSQLQuery);
//                    break;

//                case DBType.MSSQL:
//                    ReturnValue = SQLtoFloat(SQLConnection, sSQLQuery);
//                    break;

//                case DBType.Access:
//                    ReturnValue = SQLtoFloat(MSAConnection, sSQLQuery);
//                    break;

//                case DBType.Firebird:
//                    ReturnValue = SQLtoFloat(FBConnection, sSQLQuery);
//                    break;

//            }
//            DebugMsg("SQLtoFloat(" + sSQLQuery + ") = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//        /*
//        private float SQLtoFloat(OleDbConnection oSQLConnection, string sSQLQuery) //ADODB.Connection
//        {
//            ADODB.Recordset oSQLtoString_RecordSet = new ADODB.Recordset();
//            float nSQLtoFloat_Return = float.NaN;

//            try
//            {
//                if (oSQLConnection.State == 1)
//                {
//                    oSQLtoString_RecordSet.Open(sSQLQuery, oSQLConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, 0);
//                    if (!oSQLtoString_RecordSet.EOF)
//                    {
//                        if (oSQLtoString_RecordSet.Fields.Count > 0) nSQLtoFloat_Return = (float)Convert.ToDouble(oSQLtoString_RecordSet.Fields[0].Value);
//                    }
//                    oSQLtoString_RecordSet.Close();
//                }
//            }
//            catch (Exception ex)
//            {
//                if (oSQLConnection.State == 1) oSQLtoString_RecordSet.Close();
//                if (oSQLtoString_RecordSet != null) oSQLtoString_RecordSet.Close();
//                if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//            }
//            return nSQLtoFloat_Return;
//            //          return StrToFloat(SQLtoString(oSQLtoFloat_Connection, sSQLtoFloat_Query));
//        }
//        */
//        private float SQLtoFloat(OracleConnection oSQLConnection, string sSQLQuery)
//        {
//            float ReturnValue = float.NaN;
//            if (IsConnected)
//            {
//                OracleCommand Command = new OracleCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = (float)Convert.ToDouble(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private float SQLtoFloat(FbConnection oSQLConnection, string sSQLQuery)
//        {
//            float ReturnValue = float.NaN;
//            if (IsConnected)
//            {
//                FbCommand Command = new FbCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = (float)Convert.ToDouble(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }

//        private float SQLtoFloat(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            float ReturnValue = float.NaN;
//            if (IsConnected)
//            {
//                SqlCommand Command = new SqlCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = (float)Convert.ToDouble(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private float SQLtoFloat(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            float ReturnValue = float.NaN;
//            if (IsConnected)
//            {
//                OleDbCommand Command = new OleDbCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = (float)Convert.ToDouble(Command.ExecuteScalar());
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }

//        public string SQLtoString(string Field, string Table, string Where)
//        {
//            if (string.IsNullOrEmpty(Field) | string.IsNullOrEmpty(Table)) { return string.Empty; }
//            string Query = "select " + Field + " from " + Owner + Table;
//            if (!string.IsNullOrEmpty(Where)) { Query += " where " + Where; }
//            return SQLtoString(Query);
//        }
//        public string SQLtoString(string Field, string Table)
//        {
//            return SQLtoString(Field, Table, "");
//        }
//        public string SQLtoString(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            string ReturnValue = string.Empty;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    ReturnValue = SQLtoString(OraConnection, sSQLQuery);
//                    break;

//                case DBType.MSSQL:
//                    ReturnValue = SQLtoString(SQLConnection, sSQLQuery);
//                    break;

//                case DBType.Firebird:
//                    ReturnValue = SQLtoString(FBConnection, sSQLQuery);
//                    break;

//                case DBType.Access:
//                    ReturnValue = SQLtoString(MSAConnection, sSQLQuery);
//                    break;
//            }
//            DebugMsg("SQLtoString(" + sSQLQuery + ") = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//        /*
//        private string SQLtoString(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            ADODB.Recordset oSQLtoString_RecordSet = new ADODB.Recordset();
//            string sSQLtoString_Return = string.Empty;

//            try
//            {
//                if (oSQLConnection.State == 1)
//                {
//                    oSQLtoString_RecordSet.Open(sSQLQuery, oSQLConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, 0);
//                    if (!oSQLtoString_RecordSet.EOF)
//                    {
//                        if (oSQLtoString_RecordSet.Fields.Count > 0) sSQLtoString_Return = oSQLtoString_RecordSet.Fields[0].Value.ToString();
//                    }
//                    oSQLtoString_RecordSet.Close();
//                }
//            }
//            catch (Exception ex)
//            {
//                if (oSQLConnection.State == 1) oSQLtoString_RecordSet.Close();
//                if (oSQLtoString_RecordSet != null) oSQLtoString_RecordSet.Close();
//                if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//            }
//            return sSQLtoString_Return;
//        }
//        */
//        private string SQLtoString(OracleConnection oSQLConnection, string sSQLQuery)
//        {
//            string ReturnValue = string.Empty;
//            if (IsConnected)
//            {
//                OracleCommand Command = new OracleCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Command.ExecuteScalar().ToString();
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private string SQLtoString(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            string ReturnValue = string.Empty;
//            if (IsConnected)
//            {
//                SqlCommand Command = new SqlCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Command.ExecuteScalar().ToString();
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private string SQLtoString(FbConnection oSQLConnection, string sSQLQuery)
//        {
//            string ReturnValue = string.Empty;
//            if (IsConnected)
//            {
//                FbCommand Command = new FbCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Command.ExecuteScalar().ToString();
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private string SQLtoString(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            string ReturnValue = string.Empty;
//            if (IsConnected)
//            {
//                OleDbCommand Command = new OleDbCommand();
//                try
//                {
//                    Command.CommandText = sSQLQuery;
//                    Command.Connection = oSQLConnection;
//                    ReturnValue = Command.ExecuteScalar().ToString();
//                    Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (Command != null) Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
        
//        public string BLOBtoString(object LobValue)
//        {
//            return _BLOBtoString((byte[])LobValue);
//        }
//        private string _BLOBtoString(byte[] LobValue)
//        {
//            string ReturnValue = string.Empty;
//            for (long i = 0; i < LobValue.Length; i++)
//            {
//                ReturnValue += (LobValue[i]).ToString("X2");
//            }
//            return ReturnValue;
//        }
//        private string _BLOBtoString(string LobValue)
//        {
//            return LobValue;
//        }
//        private string SQLBLOBtoString2(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            string ReturnValue = string.Empty;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    //ReturnValue = SQLBLOBtoString(OraConnection, sSQLQuery);
//                    break;

//                case DBType.MSSQL:
//                    ReturnValue = SQLBLOBtoString(SQLConnection, sSQLQuery);
//                    break;

//                case DBType.Access:
//                    ReturnValue = SQLBLOBtoString(MSAConnection, sSQLQuery);
//                    break;

//            }
//            DebugMsg("SQLBLOBtoString(" + sSQLQuery + ") = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//        /*
//        private string SQLBLOBtoString(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            ADODB.Recordset oSQLtoString_RecordSet = new ADODB.Recordset();
//            string sSQLtoString_Return = string.Empty;

//            try
//            {
//                if (oSQLConnection.State == 1)
//                {
//                    oSQLtoString_RecordSet.Open(sSQLQuery, oSQLConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, 0);
//                    if (!oSQLtoString_RecordSet.EOF)
//                    {
//                        if (oSQLtoString_RecordSet.Fields.Count > 0)
//                        {
//                            for (int i = 0; i <= ((byte[])oSQLtoString_RecordSet.Fields[0].Value).GetLength(0); i += 1)
//                            {
//                                if (Convert.ToString(((byte[])oSQLtoString_RecordSet.Fields[0].Value)[i], 16).Length == 2)
//                                {
//                                    sSQLtoString_Return = sSQLtoString_Return + Convert.ToString(((byte[])oSQLtoString_RecordSet.Fields[0].Value)[i], 16);
//                                }
//                                else
//                                {
//                                    sSQLtoString_Return = sSQLtoString_Return + "0" + Convert.ToString(((byte[])oSQLtoString_RecordSet.Fields[0].Value)[i], 16);
//                                }
//                            }
//                        }
//                    }
//                    oSQLtoString_RecordSet.Close();
//                }
//            }
//            catch (Exception ex)
//            {
//                if (oSQLConnection.State == 1) oSQLtoString_RecordSet.Close();
//                if (oSQLtoString_RecordSet != null) oSQLtoString_RecordSet.Close();
//                DebugMsg("SSQLBLOBtoString(" + oSQLConnection.ConnectionString + ", " + sSQLQuery + "): " + ex.Message);
//                //if (GenericTools.Interactive) MessageBox.Show("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//            }
//            return sSQLtoString_Return.ToUpper();
//        }
//        */
//        //private string SQLBLOBtoString(OracleConnection oSQLConnection, string sSQLQuery)
//        //{
//        //    string ReturnValue = string.Empty;

//        //    if (IsConnected)
//        //    {
//        //        OracleCommand Command = new OracleCommand();
//        //        OracleDataReader Reader = null;
//        //        try
//        //        {
//        //            Command.CommandText = sSQLQuery;
//        //            Command.Connection = oSQLConnection;
//        //            Reader = Command.ExecuteReader();
//        //            if (Reader.HasRows)
//        //            {
//        //                Reader.Read();
//        //                byte[] LobValue = new byte[Convert.ToInt32(Reader.GetOracleBlob(0).Length)];
//        //                int LogSize = Reader.GetOracleBlob(0).Read(LobValue, 0, LobValue.Length);
//        //                ReturnValue = _BLOBtoString(LobValue);
//        //                Reader.Dispose();
//        //                Command.Dispose();
//        //            }
//        //        }
//        //        catch (Exception ex)
//        //        {
//        //            if (Reader != null) Reader.Dispose();
//        //            if (Command != null) Command.Dispose();
//        //            if (ex.InnerException != null) DebugMsg("LOB Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//        //        }
//        //    }
//        //    return ReturnValue;
//        //}
//        private string SQLBLOBtoString(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            return SQLtoString(oSQLConnection, sSQLQuery);
//        }
//        private string SQLBLOBtoString(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            return SQLtoString(oSQLConnection, sSQLQuery);
//        }


//        public bool SQLExec_Command(FbCommand Command)
//        {
//            DateTime StartTime = DateTime.Now;
//            bool ReturnValue = false;
//            switch (DBType)
//            {
//                case DBType.Firebird:
//                    ReturnValue = SQLExec(FBConnection, Command);
//                    break;
//            }
//            DebugMsg("SQLExec(Command) = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;


//        }
//        public bool SQLExec(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            bool ReturnValue = false;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    ReturnValue = SQLExec(OraConnection, sSQLQuery);
//                    break;

//                case DBType.MSSQL:
//                    ReturnValue = SQLExec(SQLConnection, sSQLQuery);
//                    break;

//                case DBType.Access:
//                    ReturnValue = SQLExec(MSAConnection, sSQLQuery);
//                    break;

//                case DBType.Firebird:
//                    ReturnValue = SQLExec(FBConnection, sSQLQuery);
//                    break;
//            }
//            DebugMsg("SQLExec(" + sSQLQuery + ") = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//        /*
//        private bool SQLExec(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            object Temp = new object();

//            try
//            {
//                if (oSQLConnection.State == 1)
//                {
//                    oSQLConnection.Execute(sSQLQuery, out Temp, 0);
//                    return true;
//                }
//            }
//            catch (Exception ex)
//            {
//                if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//            }
//            return false;
//        }
//        */
//        private bool SQLExec(OracleConnection oSQLConnection, string sSQLQuery)
//        {
//            bool ReturnValue = false;
//            if (IsConnected)
//            {
//                OracleCommand oSQLExec_Command = new OracleCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                    ReturnValue = true;
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private bool SQLExec(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            bool ReturnValue = false;
//            if (IsConnected)
//            {
//                SqlCommand oSQLExec_Command = new SqlCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    oSQLExec_Command.CommandTimeout = 20000;
//                    oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                    ReturnValue = true;
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }

//        private bool SQLExec(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            bool ReturnValue = false;
//            if (IsConnected)
//            {
//                OleDbCommand oSQLExec_Command = new OleDbCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                    ReturnValue = true;
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private bool SQLExec(FbConnection oSQLConnection, string sSQLQuery)
//        {
//            bool ReturnValue = false;
//            if (IsConnected)
//            {
//                FbCommand oSQLExec_Command = new FbCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                    ReturnValue = true;
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private bool SQLExec(FbConnection oSQLConnection, FbCommand Command)
//        {
//            bool ReturnValue = false;
//            if (IsConnected)
//            {
//                FbCommand oSQLExec_Command = Command;
//                try
//                {
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                    ReturnValue = true;
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine +  Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        public Int32 SQLExecInt(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            Int32 ReturnValue = 0;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    //Not implemented yet
//                    //ReturnValue = SQLExec(OraConnection, sSQLQuery);
//                    break;

//                case DBType.MSSQL:
//                    ReturnValue = SQLExecInt(SQLConnection, sSQLQuery);
//                    break;

//                case DBType.Firebird:
//                    //Not implemented yet
//                    //ReturnValue = SQLExecInt(FBConnection, sSQLQuery);
//                    break;

//                case DBType.Access:
//                    //Not implemented yet
//                    //ReturnValue = SQLExec(MSAConnection, sSQLQuery);
//                    break;
//            }
//            DebugMsg("SQLExec(" + sSQLQuery + ") = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//        private Int32 SQLExecInt(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            Int32 ReturnValue = 0;
//            if (IsConnected)
//            {
//                SqlCommand oSQLExec_Command = new SqlCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    ReturnValue = Convert.ToInt32(oSQLExec_Command.ExecuteScalar());//.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
        
//        public Int32 SQLUpdate(string Table, string Field, float NewValue, string Where)
//        {
//            if (string.IsNullOrEmpty(Field) | string.IsNullOrEmpty(Table)) { return 0; }
//            string Query = "update " + Owner + Table + " set " + Field + " = " + NewValue.ToString().Replace(",", ".");
//            if (!string.IsNullOrEmpty(Where)) { Query += " where " + Where; }
//            return SQLUpdate(Query);
//        }
//        public Int32 SQLUpdate(string Table, string Field, int NewValue, string Where)
//        {
//            if (string.IsNullOrEmpty(Field) | string.IsNullOrEmpty(Table)) { return 0; }
//            string Query = "update " + Owner + Table + " set " + Field + " = " + NewValue.ToString().Replace(",", ".");
//            if (!string.IsNullOrEmpty(Where)) { Query += " where " + Where; }
//            return SQLUpdate(Query);
//        }
//        public Int32 SQLUpdate(string Table, string Field, string NewValue, string Where)
//        {
//            if (string.IsNullOrEmpty(Field) | string.IsNullOrEmpty(Table)) { return 0; }
//            string Query = "update " + Owner + Table + " set " + Field + " = '" + NewValue + "'";
//            if (!string.IsNullOrEmpty(Where)) { Query += " where " + Where; }
//            return SQLUpdate(Query);
//        }
//        public Int32 SQLUpdate(string Table, string Field, float NewValue)
//        {
//            return SQLUpdate(Table, Field, NewValue, "");
//        }
//        public Int32 SQLUpdate(string Table, string Field, int NewValue)
//        {
//            return SQLUpdate(Table, Field, NewValue, "");
//        }
//        public Int32 SQLUpdate(string Table, string Field, string NewValue)
//        {
//            return SQLUpdate(Table, Field, NewValue, "");
//        }
//        public Int32 SQLUpdate(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            int ReturnValue = 0;
//            switch (DBType)
//            {
//                case DBType.Oracle:
//                    ReturnValue = SQLUpdate(OraConnection, sSQLQuery);
//                    break;

//                case DBType.MSSQL:
//                    ReturnValue = SQLUpdate(SQLConnection, sSQLQuery);
//                    break;

//                case DBType.Access:
//                    ReturnValue = SQLUpdate(MSAConnection, sSQLQuery);
//                    break;

//                case DBType.Firebird:
//                    ReturnValue = SQLUpdate(FBConnection, sSQLQuery);
//                    break;
//            }
//            DebugMsg("SQLUpdate(" + sSQLQuery + ") = " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//        /*
//        private Int32 SQLUpdate(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            object Temp = new object();

//            try
//            {
//                if (oSQLConnection.State == 1)
//                {
//                    oSQLConnection.Execute(sSQLQuery, out Temp, 0);
//                    return 0;
//                }
//            }
//            catch (Exception ex)
//            {
//                if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//            }
//            return 0;
//        }
//        */
//        private Int32 SQLUpdate(OracleConnection oSQLConnection, string sSQLQuery)
//        {
//            int ReturnValue = 0;
//            if (IsConnected)
//            {
//                OracleCommand oSQLExec_Command = new OracleCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    ReturnValue = oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private Int32 SQLUpdate(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            int ReturnValue = 0;
//            if (IsConnected)
//            {
//                SqlCommand oSQLExec_Command = new SqlCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    ReturnValue = oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private Int32 SQLUpdate(FbConnection oSQLConnection, string sSQLQuery)
//        {
//            int ReturnValue = 0;
//            if (IsConnected)
//            {
//                FbCommand oSQLExec_Command = new FbCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    ReturnValue = oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }
//        private Int32 SQLUpdate(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            int ReturnValue = 0;
//            if (IsConnected)
//            {
//                OleDbCommand oSQLExec_Command = new OleDbCommand();
//                try
//                {
//                    oSQLExec_Command.CommandText = sSQLQuery;
//                    oSQLExec_Command.Connection = oSQLConnection;
//                    ReturnValue = oSQLExec_Command.ExecuteNonQuery();
//                    oSQLExec_Command.Dispose();
//                }
//                catch (Exception ex)
//                {
//                    if (oSQLExec_Command != null) oSQLExec_Command.Dispose();
//                    if (ex.InnerException != null) DebugMsg("SQL Query:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return ReturnValue;
//        }

//        public int SQLInsert(string TableName, List<TableColumn> Columns, string ReturningColumnName = "", string SequenceName = "")
//        {
//            GenericTools.DebugMsg("SQLInsert(" + TableName + ", " + Columns.Count.ToString() + " columns" + (string.IsNullOrEmpty(ReturningColumnName) ? string.Empty : ", " + ReturningColumnName.ToString()) + (string.IsNullOrEmpty(SequenceName) ? string.Empty : ", " + SequenceName.ToString()) + "): Starting...");

//            try
//            {
//                switch (DBType)
//                {
//                    case DB.DBType.Oracle: return _SQLInsert(OraConnection, TableName, Columns, ReturningColumnName, SequenceName);
//                    case DB.DBType.MSSQL: return _SQLInsert(SQLConnection, TableName, Columns, ReturningColumnName);
//                    case DB.DBType.Firebird: return _SQLInsert(FBConnection, TableName, Columns, ReturningColumnName);
//                    default: return _SQLInsert(MSAConnection, TableName, Columns, ReturningColumnName);
//                }
//            }
//            catch (Exception ex)
//            {
//                GenericTools.DebugMsg("SQLInsert(" + TableName + ", " + Columns.Count.ToString() + " columns, " + ReturningColumnName.ToString() + ", " + SequenceName.ToString() + ") error: " + ex.Message);
//            }
//            return 0;
//        }
//        private int _SQLInsert(OracleConnection Connection, string TableName, List<TableColumn> Columns, string ReturningColumnName = "", string SequenceName = "")
//        {
//            int ReturnValue = 0;
//            int Id = -1;
            
//            try
//            {
//                OracleCommand Command = new OracleCommand(_SQLInsertGetString(TableName, Columns, ReturningColumnName, SequenceName), Connection); //returning CondicaoEqpId", SAMConnection); ""
                
//                foreach (TableColumn Column in Columns)
//                    if (Column.Name.ToUpper() != ReturningColumnName.ToUpper())
//                    {
//                        if (Column.Value == null)
//                            Command.Parameters.Add(":p" + Column.Name, Column.Value);
//                        else
//                            switch (Column.Value.GetType().ToString())
//                            {
//                                case "System.UInt16":
//                                    Command.Parameters.Add(":p" + Column.Name, Convert.ToInt16(Column.Value));
//                                    break;
//                                case "System.UInt32":
//                                    Command.Parameters.Add(":p" + Column.Name, Convert.ToInt32(Column.Value));
//                                    break;
//                                case "System.UInt64":
//                                    Command.Parameters.Add(":p" + Column.Name, Convert.ToInt64(Column.Value));
//                                    break;
//                                case "System.uint":
//                                    Command.Parameters.Add(":p" + Column.Name, Convert.ToInt32(Column.Value));
//                                    break;
//                                default:
//                                    Command.Parameters.Add(":p" + Column.Name, Column.Value);
//                                    break;
//                            }
//                    }
//                    else
//                    {
//                        Id = SQLtoInt("select " + Owner + SequenceName + ".NextVal from Dual");
//                        Command.Parameters.Add(":p" + Column.Name, Id);
//                    }
                
//                if (Id != 0)
//                {
//                    int Temp = Convert.ToInt32(Command.ExecuteNonQuery());
//                    ReturnValue = ((Id == -1) ? Temp : Id);
//                }
//            }
//            catch (Exception ex)
//            {
//                GenericTools.DebugMsg("_SQLInsert(" + TableName + ") error: " + ex.Message);
//            }
            
//            GenericTools.DebugMsg("_SQLInsert(" + TableName + "): " + ReturnValue.ToString());

//            return ReturnValue;
//        }
//        private int _SQLInsert(SqlConnection Connection, string TableName, List<TableColumn> Columns, string ReturningColumnName = "")
//        {
//            int ReturnValue = 0;

//            try
//            {
//                SqlCommand Command = new SqlCommand(_SQLInsertGetString(TableName, Columns, ReturningColumnName), Connection); //returning CondicaoEqpId", SAMConnection); ""
//                foreach (TableColumn Column in Columns)
//                    if (Column.Name.ToUpper() != ReturningColumnName.ToUpper())
//                    {
//                        if (Column.Value == null)
//                            Command.Parameters.AddWithValue("@" + Column.Name, Column.Value);
//                        else
//                            switch (Column.Value.GetType().ToString())
//                            {
//                                case "System.UInt16":
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Convert.ToInt16(Column.Value));
//                                    break;
//                                case "System.UInt32":
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Convert.ToInt32(Column.Value));
//                                    break;
//                                case "System.UInt64":
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Convert.ToInt64(Column.Value));
//                                    break;
//                                case "System.uint":
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Convert.ToInt32(Column.Value));
//                                    break;
//                                default:
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Column.Value);
//                                    break;
//                            }
//                    }
//                ReturnValue = Convert.ToInt32(Command.ExecuteScalar());
//            }
//            catch (Exception ex)
//            {
//                GenericTools.DebugMsg("_SQLInsert(" + TableName + ") error: " + ex.Message);
//            }
            
//            GenericTools.DebugMsg("_SQLInsert(" + TableName + "): " + ReturnValue.ToString());
            
//            return ReturnValue;
//        }
//        private int _SQLInsert(FbConnection Connection, string TableName, List<TableColumn> Columns, string ReturningColumnName = "")
//        {
//            int ReturnValue = 0;

//            try
//            {
//                FbCommand Command = new FbCommand(_SQLInsertGetString(TableName, Columns, ReturningColumnName), Connection); //returning CondicaoEqpId", SAMConnection); ""
//                foreach (TableColumn Column in Columns)
//                    if (Column.Name.ToUpper() != ReturningColumnName.ToUpper())
//                    {
//                        if (Column.Value == null)
//                            Command.Parameters.AddWithValue("@" + Column.Name, Column.Value);
//                        else
//                            switch (Column.Value.GetType().ToString())
//                            {
//                                case "System.UInt16":
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Convert.ToInt16(Column.Value));
//                                    break;
//                                case "System.UInt32":
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Convert.ToInt32(Column.Value));
//                                    break;
//                                case "System.UInt64":
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Convert.ToInt64(Column.Value));
//                                    break;
//                                case "System.uint":
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Convert.ToInt32(Column.Value));
//                                    break;
//                                default:
//                                    Command.Parameters.AddWithValue("@" + Column.Name, Column.Value);
//                                    break;
//                            }
//                    }
//                ReturnValue = Convert.ToInt32(Command.ExecuteScalar());
//            }
//            catch (Exception ex)
//            {
//                GenericTools.DebugMsg("_SQLInsert(" + TableName + ") error: " + ex.Message);
//            }

//            GenericTools.DebugMsg("_SQLInsert(" + TableName + "): " + ReturnValue.ToString());

//            return ReturnValue;
//        }
//        private int _SQLInsert(OleDbConnection Connection, string TableName, List<TableColumn> Columns, string ReturningColumnName = "")
//        {
//            int ReturnValue = 0;

//            try
//            {
//                OleDbCommand Command = new OleDbCommand(_SQLInsertGetString(TableName, Columns, ReturningColumnName), Connection); //returning CondicaoEqpId", SAMConnection); ""
//                foreach (TableColumn Column in Columns)
//                    if (Column.Name.ToUpper() != ReturningColumnName.ToUpper())
//                        Command.Parameters.AddWithValue("@" + Column.Name, Column.Value);

//                int Temp = Convert.ToInt32(Command.ExecuteScalar());
//                if (!string.IsNullOrEmpty(ReturningColumnName))
//                    ReturnValue = SQLtoInt("SELECT @@IDENTITY");
//                else
//                    ReturnValue = Temp;
//            }
//            catch (Exception ex)
//            {
//                GenericTools.DebugMsg("_SQLInsert(" + TableName + ") error: " + ex.Message);
//            }

//            GenericTools.DebugMsg("_SQLInsert(" + TableName + "): " + ReturnValue.ToString());
            
//            return ReturnValue;
//        }

//        private string _SQLInsertGetString(string TableName, List<TableColumn> Columns, string ReturningColumnName = null, string SequenceName = null)
//        {
//            StringBuilder InsertCommand = new StringBuilder();
//            StringBuilder InsertCommandValues = new StringBuilder();
//            InsertCommand.Append("insert into " + Owner + TableName + " (");
//            bool FirstColumn = true;
//            foreach (TableColumn Column in Columns)
//            {
//                if ((DBType==DBType.Oracle) || (Column.Name.ToUpper() != ReturningColumnName.ToUpper()))
//                {
//                    InsertCommand.Append((FirstColumn ? string.Empty : ",") + Column.Name);
//                    InsertCommandValues.Append((FirstColumn ? string.Empty : ",") + (DBType==DBType.Oracle ? ":p" : "@") + Column.Name);
//                    FirstColumn = false;
//                }
//            }
//            InsertCommand.Append(") values (" + InsertCommandValues.ToString() + ")");
//            if (!String.IsNullOrEmpty(ReturningColumnName))
//                switch (DBType)
//                {
//                    case DB.DBType.MSSQL: return (InsertCommand.ToString() + ";SELECT SCOPE_IDENTITY();");
//                    case DB.DBType.Firebird: return (InsertCommand.ToString() + " returning " + ReturningColumnName);
//                    default: return InsertCommand.ToString();
//                }

//            return InsertCommand.ToString();
//        }

//        public DataTable RecordSet(string sSQLQuery) { return DataTable(sSQLQuery); }
//        public DataTable RecordSet(string Field, string Table) { return DataTable(Field, Table, ""); }
//        public DataTable RecordSet(string Field, string Table, string Where) { return DataTable(Field, Table, Where); }
//        public DataTable DataTable(string Field, string Table) { return DataTable(Field, Table, ""); }
//        public DataTable DataTable(string Field, string Table, string Where)
//        {
//            if (string.IsNullOrEmpty(Field) | string.IsNullOrEmpty(Table)) { return null; }
//            string Query = "select " + Field + " from " + Owner + Table;
//            if (!string.IsNullOrEmpty(Where)) { Query += " where " + Where; }
//            return DataTable(Query);
//        }
//        public DataTable DataTable(string sSQLQuery)
//        {
//            DateTime StartTime = DateTime.Now;
//            DebugMsg("RecordSet(" + sSQLQuery + ") Starting... - " + (DateTime.Now - StartTime).ToString());
//            DataTable oRecordSet_DataTable = new DataTable();
//            try
//            {
//                if (IsConnected)
//                {
//                    switch (DBType)
//                    {
//                        case DBType.Oracle:
//                            oRecordSet_DataTable = DataTable(OraConnection, sSQLQuery);
//                            break;

//                        case DBType.MSSQL:
//                            oRecordSet_DataTable = DataTable(SQLConnection, sSQLQuery);
//                            break;

//                        case DBType.Access:
//                            oRecordSet_DataTable = DataTable(MSAConnection, sSQLQuery);
//                            break;

//                        case DBType.Firebird:
//                            oRecordSet_DataTable = DataTable(FBConnection, sSQLQuery);
//                            break;
//                    }
//                }
//                else
//                {
//                    DebugMsg("RecordSet(" + sSQLQuery + ") not connected to DB - " + (DateTime.Now - StartTime).ToString());
//                }
//            }
//            catch (Exception ex)
//            {
//                DebugMsg("RecordSet(" + sSQLQuery + ") error: " + ex.Message);
//            }
//            DebugMsg("RecordSet(" + sSQLQuery + ") = Table with " + oRecordSet_DataTable.Rows.Count.ToString() + " rows and " + oRecordSet_DataTable.Columns.Count.ToString() + " columns - " + (DateTime.Now - StartTime).ToString());
//            return oRecordSet_DataTable;
//        }
//        private DataTable DataTable(OracleConnection oSQLConnection, string sSQLQuery)
//        {
//            DataTable oRecordSet_DataTable = new DataTable();
//            if (IsConnected)
//            {
//                try
//                {

//                    OracleDataAdapter oRecordSet_DataAdapter = new OracleDataAdapter(sSQLQuery, oSQLConnection);
//                    try
//                    {
//                        oRecordSet_DataAdapter.Fill(oRecordSet_DataTable);
//                        oRecordSet_DataAdapter.Dispose();
//                    }
//                    catch (Exception ex)
//                    {
//                        if (oRecordSet_DataAdapter != null) oRecordSet_DataAdapter.Dispose();
//                        //if (ex.InnerException != null) 
//                            DebugMsg("RecordSet error: " + ex.Message);
//                    }
//                }
//                catch (Exception ex)
//                {
//                    DebugMsg("RecordSet error: " + ex.Message);
//                }
//            }
//            return oRecordSet_DataTable;
//        }
//        private DataTable DataTable(SqlConnection oSQLConnection, string sSQLQuery)
//        {
//            DataTable oRecordSet_DataTable = new DataTable();
//            if (IsConnected)
//            {
//                try
//                {
//                    SqlDataAdapter oRecordSet_DataAdapter = new SqlDataAdapter(sSQLQuery, oSQLConnection);
//                    try
//                    {
//                        oRecordSet_DataAdapter.SelectCommand.CommandTimeout = 20000;
//                        oRecordSet_DataAdapter.Fill(oRecordSet_DataTable);
//                        oRecordSet_DataAdapter.Dispose();
//                    }
//                    catch (Exception ex)
//                    {
//                        if (oRecordSet_DataAdapter != null) oRecordSet_DataAdapter.Dispose();
//                        if (ex.InnerException != null) DebugMsg("RecordSet:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                    }
//                }
//                catch (Exception ex)
//                {
//                    DebugMsg("RecordSet:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return oRecordSet_DataTable;
//        }
//        private DataTable DataTable(OleDbConnection oSQLConnection, string sSQLQuery)
//        {
//            DataTable oRecordSet_DataTable = new DataTable();
//            if (IsConnected)
//            {
//                try
//                {

//                    OleDbDataAdapter oRecordSet_DataAdapter = new OleDbDataAdapter(sSQLQuery, oSQLConnection);
//                    try
//                    {
//                        oRecordSet_DataAdapter.Fill(oRecordSet_DataTable);
//                        oRecordSet_DataAdapter.Dispose();
//                    }
//                    catch (Exception ex)
//                    {
//                        if (oRecordSet_DataAdapter != null) oRecordSet_DataAdapter.Dispose();
//                        if (ex.InnerException != null) DebugMsg("RecordSet:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                    }
//                }
//                catch (Exception ex)
//                {
//                    DebugMsg("RecordSet:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return oRecordSet_DataTable;
//        }
//        private DataTable DataTable(FbConnection oSQLConnection, string sSQLQuery)
//        {
//            DataTable oRecordSet_DataTable = new DataTable();
//            if (IsConnected)
//            {
//                try
//                {

//                    FbDataAdapter oRecordSet_DataAdapter = new FbDataAdapter(sSQLQuery, oSQLConnection);
//                    try
//                    {
//                        oRecordSet_DataAdapter.Fill(oRecordSet_DataTable);
//                        oRecordSet_DataAdapter.Dispose();
//                    }
//                    catch (Exception ex)
//                    {
//                        if (oRecordSet_DataAdapter != null) oRecordSet_DataAdapter.Dispose();
//                        //if (ex.InnerException != null) 
//                        DebugMsg("RecordSet:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                    }
//                }
//                catch (Exception ex)
//                {
//                    DebugMsg("RecordSet:" + Environment.NewLine + sSQLQuery + Environment.NewLine + ex.Message);
//                }
//            }
//            return oRecordSet_DataTable;
//        }

//        public bool Connect(string sConnectionString, string DbUser, string DbPassword, DBType nDBType)
//        {
//            User = DbUser;
//            Password = DbPassword;
//            return Connect(sConnectionString + ";User ID='" + User + "';Password='" + Password + "'", nDBType);
//        }
//        public bool Connect(string sConnectionString, DBType nDBType)
//        {
//            ConnectionString = sConnectionString;
//            DBType = nDBType;
//            return Connect();
//        }
//        public bool Connect()
//        {
//            GenericTools.DebugMsg("Connect(): Starting...");
//            GenericTools.LastError.Clear();
//            string sTempConnectionString = ConnectionString;

//            if (String.IsNullOrEmpty(ConnectionString))
//            {
//                if (String.IsNullOrEmpty(DBName))
//                {
//                    GenericTools.DebugMsg("Connect(" + ConnectionString + "): Empty DBName");
//                    ConnectionString = sTempConnectionString;
//                    return false;
//                }
//                GenericTools.DebugMsg("Connect(): DBType=" + DBType.ToString());
//                switch (DBType)
//                {
//                    case DBType.Oracle:
//                        ConnectionString = "User Id=" + User + ";Password=" + Password + ";Data Source=" + DBName;
//                        Owner = "SKFUser1.";
//                        break;

//                    case DBType.MSSQL:
//                        if (String.IsNullOrEmpty(InitialCatalog))
//                        {
//                            GenericTools.LastError.Message = "DBConnection.Open(" + ConnectionString + "): Empty InitialCatalog" + Environment.NewLine + this.ToString();
//                            ConnectionString = sTempConnectionString;
//                            return false;
//                        }
//                        //ConnectionString = "Data source='" + DBName + "';Initial Catalog=skfuser;User Id='" + User + "';Password='" + Password + "'";
//                        ConnectionString = String.Format("Data source='{0}';Initial Catalog='{1}';User Id='{2}';Password='{3}'", DBName, InitialCatalog ,User, Password);
//                        Owner = "[SKFUser1].";
//                        break;

//                    case DBType.Access:
//                        ConnectionString = "Provider='Microsoft.Jet.OLEDB.4.0';Data source='" + DBName + "';";
//                        Owner = "";
//                        break;

//                    default:
//                        GenericTools.DebugMsg("Connect(" + ConnectionString + "): Invalid DBType (" + DBType.ToString() + ")");
//                        ConnectionString = sTempConnectionString;
//                        return false;
//                }
//            }
//            GenericTools.DebugMsg("Connect(" + ConnectionString + ")");
//            try
//            {
//                switch (DBType)
//                {
//                    case DBType.Oracle:
//                        OraConnection = new OracleConnection(ConnectionString);
//                        OraConnection.Open();
//                        _IsConnected = (OraConnection.State == ConnectionState.Open);
//                        break;

//                    case DBType.MSSQL:
//                        SQLConnection.ConnectionString = ConnectionString;
//                        SQLConnection.Open();
//                        _IsConnected = (SQLConnection.State == ConnectionState.Open);
//                        break;

//                    case DBType.Access:
//                        MSAConnection.ConnectionString = ConnectionString;
//                        MSAConnection.Open();
//                        _IsConnected = (MSAConnection.State == ConnectionState.Open);
//                        break;

//                    case DBType.Firebird:
//                        FBConnection.ConnectionString = ConnectionString;
//                        FBConnection.Open();
//                        _IsConnected = (FBConnection.State == ConnectionState.Open);
//                        break;

//                    default:
//                        GenericTools.DebugMsg("Connect(" + ConnectionString + "): Invalid DBType (" + DBType.ToString() + ")");
//                        ConnectionString = sTempConnectionString;
//                        return false;
//                }
//            }
//            catch (Exception ex)
//            {
                
//                GenericTools.DebugMsg("Connect(" + ConnectionString + "): DB Open: " + ex.Message);
//                ConnectionString = sTempConnectionString;
//                return false;
//            }
//            return IsConnected;
//        }

//        public bool Close()
//        {

//            return Close(true);
//        }
//        public bool Close(bool ClearConnectionString)
//        {
//            GenericTools.DebugMsg("Connect(" + ConnectionString + "): Starting...");
//            try
//            {
//                if (IsConnected)
//                {
//                    if (DBType == DBType.Oracle) OraConnection.Close();
//                    if (DBType == DBType.MSSQL) SQLConnection.Close();
//                    if (DBType == DBType.Access) MSAConnection.Close();
//                    if (DBType == DBType.Firebird) FBConnection.Close();
//                }
//                _IsConnected = false;
//                if (ClearConnectionString) ConnectionString = string.Empty;
//            }
//            catch (Exception ex)
//            {
//                GenericTools.DebugMsg("Close(" + ConnectionString + "): DB Close: " + ex.Message);
//            }
//            GenericTools.DebugMsg("Close(" + ConnectionString + "): " + (!IsConnected).ToString());
//            return (!IsConnected);
//        }

//        public float[] GetBLOBtoArray(object LobValue, float FFTFactor)
//        {
//            float[] ReturnValue = new float[0];

//            switch (DBType)
//            {
//                case DBType.Oracle:
//                case DBType.MSSQL:
//                    ReturnValue = _GetBLOBtoArray((byte[])LobValue, FFTFactor);
//                    break;

//                default:
//                    ReturnValue = _GetBLOBtoArray(LobValue.ToString(), FFTFactor);
//                    break;
//            }
//            return ReturnValue;
//        }
//        private float[] _GetBLOBtoArray(byte[] LobValue, float FFTFactor)
//        {
//            DateTime StartTime = DateTime.Now;
//            byte[] LobArray = new byte[LobValue.Length];

//            float[] ReturnValue = new float[Convert.ToInt32(LobValue.Length / 2)];

//            int FirstByte;
//            int SecondByte;

//            for (int i = 0; i < LobValue.Length; i = i + 2)
//            {
//                FirstByte = LobValue[i];//.ReadByte();
//                SecondByte = LobValue[i + 1];//.ReadByte();
//                if (SecondByte > 127)
//                {
//                    SecondByte = SecondByte - 256;
//                }
//                ReturnValue[Convert.ToInt32(i / 2)] = ((SecondByte * 256) + FirstByte) * FFTFactor;
//            }
//            DebugMsg("_GetBLOBtoArray(byte[]): " + ReturnValue.Length.ToString() + " values - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//        private float[] _GetBLOBtoArray(string LobValue, float FFTFactor)
//        {
//            DateTime StartTime = DateTime.Now;
//            byte[] LobArray = new byte[LobValue.Length];
//            float[] ReturnValue = new float[Convert.ToInt32(LobValue.Length / 2)];
//            for (int i = 0; i <= LobValue.Length; i = i + 2)
//            {
//                ReturnValue[Convert.ToInt32(i / 2)] = GenericTools.Hex2ToValue(LobValue.Substring(i, 2)) * FFTFactor;
//            }
//            DebugMsg("_GetBLOBtoArray(string): " + ReturnValue.Length.ToString() + " values - " + (DateTime.Now - StartTime).ToString());
//            return ReturnValue;
//        }
//    //    private float[] _GetBLOBtoArray2Old(Int32 nMeasId, float FFTFactor)
//    //    {
//    //        if (DBType != DBType.Oracle) return null;

//    //        Oracle.DataAccess.Types.OracleBlob LobValue;// = null; // Oracle.DataAccess.Types.OracleBlob.Null;
//    //        float[] ReturnValue = new float[0];

//    //        if (IsConnected)
//    //        {
//    //            try
//    //            {
//    //                OracleCommand Command = new OracleCommand("select ReadingData from " + Owner + "MeasReading where ReadingType=(select RegistrationId from " + Owner + "Registration where Signature='SKFCM_ASMD_FFT') and MeasId=" + nMeasId.ToString(), OraConnection);
//    //                try
//    //                {
//    //                    OracleDataReader Reader = Command.ExecuteReader();
//    //                    try
//    //                    {
//    //                        if (Reader.HasRows)
//    //                        {
//    //                            Reader.Read();
//    //                            LobValue = Reader.GetOracleBlob(0);

//    //                            Array.Resize<float>(ref ReturnValue, Convert.ToInt32(LobValue.Length / 2));

//    //                            int FirstByte;
//    //                            int SecondByte;

//    //                            for (int i = 0; i <= LobValue.Length; i = i + 2)
//    //                            {
//    //                                FirstByte = LobValue.ReadByte();
//    //                                SecondByte = LobValue.ReadByte();
//    //                                if (SecondByte > 127)
//    //                                {
//    //                                    SecondByte = SecondByte - 256;
//    //                                }
//    //                                ReturnValue[Convert.ToInt32(i / 2)] = ((SecondByte * 256) + FirstByte) * FFTFactor;
//    //                            }

//    //                            Reader.Dispose();
//    //                        }
//    //                    }
//    //                    catch (Exception ex)
//    //                    {
//    //                        if (Reader != null) Reader.Dispose();
//    //                        if (ex.InnerException != null) DebugMsg("GetBLOBtoArray(" + nMeasId + "):" + ex.Message);
//    //                    }
//    //                }
//    //                catch (Exception ex)
//    //                {
//    //                    if (Command != null) Command.Dispose();
//    //                    if (ex.InnerException != null) DebugMsg("GetBLOBtoArray(" + nMeasId + "):" + ex.Message);
//    //                }
//    //                Command.Dispose();
//    //            }
//    //            catch (Exception ex)
//    //            {
//    //                if (ex.InnerException != null) DebugMsg("GetBLOBtoArray(" + nMeasId + "):" + ex.Message);
//    //            }
//    //        }
//    //        return ReturnValue;
//    //    }
//    //    private float[] _GetBLOBtoArray2(Int32 nMeasId, float FFTFactor)
//    //    {
//    //        DateTime StartTime = DateTime.Now;

//    //        Oracle.DataAccess.Types.OracleBlob LobValue;// = null;// Oracle.DataAccess.Types.OracleBlob.Null;
//    //        float[] ReturnValue = new float[0];

//    //        if (DBType == DBType.Oracle & IsConnected)
//    //        {
//    //            try
//    //            {
//    //                OracleCommand Command = new OracleCommand("select ReadingData from " + Owner + "MeasReading where ReadingType==(select RegistrationId from " + Owner + "Registration where Signature='SKFCM_ASMD_FFT') and MeasId=" + nMeasId.ToString(), OraConnection);
//    //                try
//    //                {
//    //                    OracleDataReader Reader = Command.ExecuteReader();
//    //                    try
//    //                    {
//    //                        if (Reader.HasRows)
//    //                        {
//    //                            Reader.Read();
//    //                            LobValue = Reader.GetOracleBlob(0);
//    //                            byte[] LobArray = new byte[LobValue.Length];

//    //                            long LobSize = LobValue.Read(LobArray, 0, (int)LobValue.Length);

//    //                            Array.Resize<float>(ref ReturnValue, Convert.ToInt32(LobSize / 2));

//    //                            int FirstByte;
//    //                            int SecondByte;

//    //                            for (int i = 0; i <= LobSize; i = i + 2)
//    //                            {
//    //                                FirstByte = LobArray[i];//.ReadByte();
//    //                                SecondByte = LobArray[i + 1];//.ReadByte();
//    //                                if (SecondByte > 127)
//    //                                {
//    //                                    SecondByte = SecondByte - 256;
//    //                                }
//    //                                ReturnValue[Convert.ToInt32(i / 2)] = ((SecondByte * 256) + FirstByte) * FFTFactor;
//    //                            }

//    //                            Reader.Dispose();
//    //                        }
//    //                    }
//    //                    catch (Exception ex)
//    //                    {
//    //                        if (Reader != null) Reader.Dispose();
//    //                        if (ex.InnerException != null) DebugMsg("GetBLOBtoArray(" + nMeasId + "):" + ex.Message);
//    //                    }
//    //                }
//    //                catch (Exception ex)
//    //                {
//    //                    if (Command != null) Command.Dispose();
//    //                    if (ex.InnerException != null) DebugMsg("GetBLOBtoArray(" + nMeasId + "):" + ex.Message);
//    //                }
//    //                Command.Dispose();
//    //            }
//    //            catch (Exception ex)
//    //            {
//    //                if (ex.InnerException != null) DebugMsg("GetBLOBtoArray(" + nMeasId + "):" + ex.Message);
//    //            }
//    //        }
//    //        DebugMsg("GetBLOBtoArray(" + nMeasId.ToString() + "): " + ReturnValue.Length.ToString() + " values - " + (DateTime.Now - StartTime).ToString());
//    //        return ReturnValue;
//    //    }

       
//        public bool CopyTable(string FromTable, string DestTable = "")
//        {
//            bool _CopyTable = false;

//            switch (DBType)
//            {
//                case DB.DBType.Oracle:

//                    break;

//                case DB.DBType.MSSQL:
//                    break;
//            }
//            //Verificar ase tabelas exeitem
//            //create 


//            return _CopyTable;
//        }

//    }
//}
