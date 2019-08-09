using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Net;
using System.IO;

using SKF.RS.STB.DB;
using SKF.RS.STB.Generic;

using SOAPCon;
namespace SKF.RS.STB.Analyst
{
    public class MeasurementDeep
    {
        private AnalystConnection Conex { get; set; }
        public MeasurementDeep(AnalystConnection _Conex)
        {
            this.Conex = _Conex;
        }

        public spectrumHeader getHeadReading_FFT(int MeasId)
        {
            return readHeadMeas(getByteHeadReading_FFT(MeasId));
        }
        public double[] getDataReading_FFT(int MeasId)
        {
            spectrumHeader header = getHeadReading_FFT(MeasId);
            return readReadingMeas(getByteDataReading_FFT(MeasId), header);
        
        }

        public spectrumHeader getHeadReading_TIME(int MeasId)
        {
            return readHeadMeas(getByteHeadReading_Time(MeasId));
        }
        public double[] getDataReading_TIME(int MeasId)
        {
            spectrumHeader header = getHeadReading_TIME(MeasId);
            return readReadingMeas(getByteDataReading_Time(MeasId), header);
        }

        public int storeData_FFT(int PointId, DateTime Timestamp, spectrumHeader Head, double[] Data, bool New_Measurement = true, int MeasId = int.MinValue) { return storeMeasurementData(PointId, Timestamp, "SKFCM_ASMD_FFT", Head, Data, New_Measurement, MeasId); }
        public int storeData_Time(int PointId, DateTime Timestamp, spectrumHeader Head, double[] Data, bool New_Measurement = true, int MeasId = int.MinValue) { return storeMeasurementData(PointId, Timestamp, "SKFCM_ASMD_Time", Head, Data, New_Measurement, MeasId); }

        private int storeMeasurementData(int PointId, DateTime Timestamp, string TypeOfData, spectrumHeader Head, double[] Data, bool New_Measurement, int MeasId)
        {
            int _return = int.MinValue;
            int _regId = (int)Registration.RegistrationId(Conex, TypeOfData);

            string _Timestamp = Generic.GenericTools.DateTime(Timestamp);

            double Fator;
            if (TypeOfData == "SKFCM_ASMD_Time")
            {
                double maxValue = double.MinValue;
                foreach (double value in Data)
                {
                    if (Math.Abs(value) > maxValue)
                        maxValue = Math.Abs(value);
                } 
                Fator = maxValue / (double)short.MaxValue;
                Head.Factor = Fator; 
            }
            else
            {
                Calculate_Factor(Data, out Fator);
                Head.Factor = Fator; 
            }

            byte[] storingHead = writeHeadMeas(Head);
            byte[] storingData = writeDataMeas(Data, Head);

            _return = storeMeasurementDB(PointId, _Timestamp, _regId, storingHead, storingData, New_Measurement, MeasId);

            return _return;
       }

       private void Calculate_Factor(double[] iData, out double oFactor )
       {
            const int WAVEFORM_PEAK_INTEGER = Int16.MaxValue - 1;
            double maxValue = double.MinValue;

            foreach ( double value in iData )
            {
            if( Math.Abs( value ) > maxValue )
                maxValue = Math.Abs( value );
            }

            oFactor = maxValue / WAVEFORM_PEAK_INTEGER;
       }
       private int storeMeasurementDB(int PointId, string Timestamp, int type, byte[] dataHeader, byte[] dataReading, bool New_Measurement, int MeasId)
        {
            
            int _return = int.MinValue;
            if (New_Measurement)
            {
                string sql_measurement = "INSERT INTO [skfuser1].[MEASUREMENT]([POINTID],[ALARMLEVEL],[STATUS],[INCLUDE],[OPERNAME],[SERIALNO],[DATADTG],[MEASUREMENTTYPE]) VALUES (" + PointId + ",0,1,1,'SKF STB Tool','STB','" + Timestamp + "',0);select @@Identity;";
                int sql_measurement_return = Conex.SQLtoInt(sql_measurement);
            }

            int max_meas = MeasId;
            if (max_meas == int.MinValue)
            {
                max_meas = Conex.SQLtoInt("SELECT MAX(MEASID) FROM MEASUREMENT");
            }

            string sql_measreading  = "INSERT INTO [skfuser1].[MEASREADING] ([MEASID],[POINTID],[READINGTYPE],[CHANNEL],[OVERALLVALUE],[READINGHEADER],[READINGDATA],[EXDWORDVAL1],[EXDOUBLEVAL1],[EXDOUBLEVAL2],[EXDOUBLEVAL3]) VALUES (" +
                                      max_meas + "," +
                                      PointId + ", " +
                                      type + "," +
                                      "1,0,@READINGHEADER,@READINGDATA,0,0,0,0);";
            
            
            using(SqlConnection _con = new SqlConnection(Conex.ConnectionString))
            using(SqlCommand _cmd = new SqlCommand(sql_measreading, _con))
            {
                SqlParameter param = _cmd.Parameters.Add("@READINGHEADER", SqlDbType.VarBinary);
                SqlParameter param2 = _cmd.Parameters.Add("@READINGDATA", SqlDbType.VarBinary);
                param.Value = dataHeader;
                param2.Value = dataReading;

                _con.Open();
                _cmd.ExecuteNonQuery();
                _con.Close();
            }
            _return = Conex.SQLtoInt("SELECT MAX(MEASID) FROM MEASREADING");
            return _return;
        }

        private byte[] getByteHeadReading_FFT(int MeasId) { return getInfo(MeasId, "READINGHEADER", "SKFCM_ASMD_FFT"); }
        private byte[] getByteDataReading_FFT(int MeasId) { return getInfo(MeasId, "READINGDATA", "SKFCM_ASMD_FFT"); }
        private byte[] getByteHeadReading_Time(int MeasId) { return getInfo(MeasId, "READINGHEADER", "SKFCM_ASMD_Time"); }
        private byte[] getByteDataReading_Time(int MeasId) { return getInfo(MeasId, "READINGDATA", "SKFCM_ASMD_Time"); }

        private byte[] getInfo(int MeasId, string Field, string type)
        {
            uint _regId = Registration.RegistrationId(Conex, type);
            byte[] _return = getByte(Field, MeasId, (int)_regId);

            return _return;
        }

        private const int mByteBoundaryOffset = 6;

        private spectrumHeader readHeadMeas(byte[] blobFile)
        {
            byte[] mHeader = blobFile;

            MemoryStream headerStream = new MemoryStream();
            headerStream.Write(mHeader, 0, mHeader.Length);
            headerStream.Seek(0, SeekOrigin.Begin);


            BinaryReader reader = new BinaryReader(headerStream);

            byte[] trash;

            spectrumHeader mSpectrumHeader = new spectrumHeader();
            mSpectrumHeader.Version = reader.ReadUInt16();

            trash = reader.ReadBytes(mByteBoundaryOffset);

            mSpectrumHeader.NumberOfLines = reader.ReadUInt16();
            mSpectrumHeader.NumberOfAverages = reader.ReadUInt16();
            mSpectrumHeader.Windowing = reader.ReadUInt32();
            mSpectrumHeader.MotorSpeed = reader.ReadDouble();
            mSpectrumHeader.Factor = reader.ReadDouble();
            mSpectrumHeader.Offset = reader.ReadDouble();
            mSpectrumHeader.HighFrequency = reader.ReadDouble();
            mSpectrumHeader.LowFrequency = reader.ReadDouble();
            mSpectrumHeader.Load = reader.ReadDouble();
            mSpectrumHeader.Detection = reader.ReadUInt32();
            mSpectrumHeader.Final = reader.ReadUInt32();
            mSpectrumHeader.Data = mHeader;

            return mSpectrumHeader;
        }
        private byte[] writeHeadMeas(spectrumHeader mSpectrumHeader)
        {

            MemoryStream headerWrite = new MemoryStream();

            byte[] trash = { 0, 0, 0, 0, 0, 0 };

            BinaryWriter writer = new BinaryWriter(headerWrite);
            writer.Write(mSpectrumHeader.Version);
            writer.Write(trash);
            writer.Write(mSpectrumHeader.NumberOfLines);
            writer.Write(mSpectrumHeader.NumberOfAverages);
            writer.Write(mSpectrumHeader.Windowing);
            writer.Write(mSpectrumHeader.MotorSpeed);
            writer.Write(mSpectrumHeader.Factor);
            writer.Write(mSpectrumHeader.Offset);
            writer.Write(mSpectrumHeader.HighFrequency);
            writer.Write(mSpectrumHeader.LowFrequency);
            writer.Write(mSpectrumHeader.Load);
            writer.Write(mSpectrumHeader.Detection);
            writer.Write(mSpectrumHeader.Final);

            byte[] write_result = headerWrite.ToArray();

            return write_result;
        }

        private double[] readReadingMeas(byte[] blobFile, spectrumHeader mSpectrumHeader)
        {
            byte[] mData = blobFile;
            double[] BinValues;
            short[] scaleValues = new short[mData.Length / sizeof(short)];
            Buffer.BlockCopy(mData, 0, scaleValues, 0, mData.Length);

            BinValues = new double[scaleValues.Length];

            for (int i = 0; i < scaleValues.Length; i++)
            {
                BinValues[i] = RestoreValue(scaleValues[i], mSpectrumHeader);
            }
            return BinValues;
        }
        private byte[] writeDataMeas(double[] BinValues, spectrumHeader mSpectrumHeader, bool _ConvertValue = true)
        {
            short[] BinValues_retornando = new short[BinValues.Length];
            byte[] retornando_para_byte = new byte[BinValues.Length * sizeof(short)];
            if (_ConvertValue)
            {
                for (int i = 0; i < BinValues.Length; i++)
                {
                    BinValues_retornando[i] = ConvertValue(BinValues[i], mSpectrumHeader);
                }

                Buffer.BlockCopy(BinValues_retornando, 0, retornando_para_byte, 0, BinValues.Length * sizeof(short));
            }
            else
            {
                Buffer.BlockCopy(BinValues, 0, retornando_para_byte, 0, BinValues.Length * sizeof(short));
            }
            return retornando_para_byte;
        }

        private double RestoreValue(short iInput, spectrumHeader mSpectrumHeader)
        {
            return (double)(iInput * mSpectrumHeader.Factor);
        }
        private short ConvertValue(double iInput, spectrumHeader mSpectrumHeader)
        {
            return (short)(iInput / mSpectrumHeader.Factor);
        }

        private byte[] getByte(string Field, int MeasId, int Type = int.MinValue)
        {
            SqlConnection connection = new SqlConnection(Conex.ConnectionString);
            connection.Open();
            string sql = "SELECT " + Field + " FROM MEASREADING WHERE MEASID=" + MeasId + " AND READINGTYPE=" + Type;
            SqlCommand command = new SqlCommand(sql, connection);
            byte[] buffer = (byte[])command.ExecuteScalar();
            connection.Close();
            return buffer;

       }

    }
    public class MeasurementDeep_old
    {
        private AnalystConnection Conex { get; set; }
        public MeasurementDeep_old(AnalystConnection _Conex)
        {
            this.Conex = _Conex;
        }

        public spectrumHeader getHeadReading_FFT(int MeasId)
        {
            return readHeadMeas(getByteHeadReading_FFT(MeasId));
        }
        public double[] getDataReading_FFT(int MeasId)
        {
            spectrumHeader header = getHeadReading_FFT(MeasId);
            return readReadingMeas(getByteDataReading_FFT(MeasId), header);
        
        }

        public spectrumHeader getHeadReading_TIME(int MeasId)
        {
            return readHeadMeas(getByteHeadReading_Time(MeasId));
        }
        public double[] getDataReading_TIME(int MeasId)
        {
            spectrumHeader header = getHeadReading_TIME(MeasId);
            return readReadingMeas(getByteDataReading_Time(MeasId), header);
        }

        public int storeData_FFT(int PointId, DateTime Timestamp, spectrumHeader Head, double[] Data, bool New_Measurement = true, int MeasId = int.MinValue) { return storeMeasurementData(PointId, Timestamp, "SKFCM_ASMD_FFT", Head, Data, New_Measurement, MeasId); }
        public int storeData_Time(int PointId, DateTime Timestamp, spectrumHeader Head, double[] Data, bool New_Measurement = true, int MeasId = int.MinValue) { return storeMeasurementData(PointId, Timestamp, "SKFCM_ASMD_Time", Head, Data, New_Measurement, MeasId); }

        private int storeMeasurementData(int PointId, DateTime Timestamp, string TypeOfData, spectrumHeader Head, double[] Data, bool New_Measurement, int MeasId)
        {
            
            int _return = int.MinValue;
            int _regId = (int)Registration.RegistrationId(Conex, TypeOfData);
            try
            {

                string _Timestamp = Generic.GenericTools.DateTime(Timestamp);

                double Fator;
                byte[] storingHead = null;
                byte[] storingData = null;

                
                    if (TypeOfData == "SKFCM_ASMD_Time")
                    {
                        if (Head.Factor == 0)
                        {
                            double maxValue = double.MinValue;
                            foreach (double value in Data)
                            {
                                if (Math.Abs(value) > maxValue)
                                    maxValue = Math.Abs(value);
                            }
                            Fator = maxValue / (double)short.MaxValue;
                            Head.Factor = Fator;
                        }
                        storingHead = writeHeadMeas(Head);
                        storingData = writeDataMeas(Data, Head);
                    }
                    else
                    {
                        if (Head.Factor == 0)
                        {
                            Calculate_Factor(Data, out Fator);
                            Head.Factor = Fator;
                        }

                        storingHead = writeHeadMeas(Head);
                        storingData = writeDataMeas(Data, Head);
                    }

                _return = storeMeasurementDB(PointId, _Timestamp, _regId, storingHead, storingData, New_Measurement, MeasId);
            }
            catch(Exception ex)
            {
                GenericTools.WriteLog("storeMeasurementData - Erro ao gravar espectro: " + ex.StackTrace);
            }
            return _return;
       }

       public void Calculate_Factor(double[] iData, out double oFactor )
       {
            const int WAVEFORM_PEAK_INTEGER = Int16.MaxValue - 1;
            double maxValue = double.MinValue;

            foreach ( double value in iData )
            {
            if( Math.Abs( value ) > maxValue )
                maxValue = Math.Abs( value );
            }

            oFactor = maxValue / WAVEFORM_PEAK_INTEGER;
       }
       private int storeMeasurementDB(int PointId, string Timestamp, int type, byte[] dataHeader, byte[] dataReading, bool New_Measurement, int MeasId)
        {
            
            int _return = int.MinValue;
            if (New_Measurement)
            {
                string sql_measurement = "INSERT INTO [skfuser1].[MEASUREMENT]([POINTID],[ALARMLEVEL],[STATUS],[INCLUDE],[OPERNAME],[SERIALNO],[DATADTG],[MEASUREMENTTYPE]) VALUES (" + PointId + ",0,1,1,'SKF STB Tool','STB','" + Timestamp + "',0);select @@Identity;";
                int sql_measurement_return = Conex.SQLtoInt(sql_measurement);
            }

            int max_meas = MeasId;
            if (max_meas == int.MinValue)
            {
                max_meas = Conex.SQLtoInt("SELECT MAX(MEASID) FROM MEASUREMENT");
            }

            string sql_measreading  = "INSERT INTO [skfuser1].[MEASREADING] ([MEASID],[POINTID],[READINGTYPE],[CHANNEL],[OVERALLVALUE],[READINGHEADER],[READINGDATA],[EXDWORDVAL1],[EXDOUBLEVAL1],[EXDOUBLEVAL2],[EXDOUBLEVAL3]) VALUES (" +
                                      max_meas + "," +
                                      PointId + ", " +
                                      type + "," +
                                      "1,0,@READINGHEADER,@READINGDATA,0,0,0,0);";
            
            
            using(SqlConnection _con = new SqlConnection(Conex.ConnectionString))
            using(SqlCommand _cmd = new SqlCommand(sql_measreading, _con))
            {
                SqlParameter param = _cmd.Parameters.Add("@READINGHEADER", SqlDbType.VarBinary);
                SqlParameter param2 = _cmd.Parameters.Add("@READINGDATA", SqlDbType.VarBinary);
                param.Value = dataHeader;
                param2.Value = dataReading;

                _con.Open();
                _cmd.ExecuteNonQuery();
                _con.Close();
            }
            _return = Conex.SQLtoInt("SELECT MAX(MEASID) FROM MEASREADING");
            return _return;
        }

        private byte[] getByteHeadReading_FFT(int MeasId) { return getInfo(MeasId, "READINGHEADER", "SKFCM_ASMD_FFT"); }
        private byte[] getByteDataReading_FFT(int MeasId) { return getInfo(MeasId, "READINGDATA", "SKFCM_ASMD_FFT"); }
        private byte[] getByteHeadReading_Time(int MeasId) { return getInfo(MeasId, "READINGHEADER", "SKFCM_ASMD_Time"); }
        private byte[] getByteDataReading_Time(int MeasId) { return getInfo(MeasId, "READINGDATA", "SKFCM_ASMD_Time"); }

        private byte[] getInfo(int MeasId, string Field, string type)
        {
            uint _regId = Registration.RegistrationId(Conex, type);
            byte[] _return = getByte(Field, MeasId, (int)_regId);

            return _return;
        }

        private const int mByteBoundaryOffset = 6;

        private spectrumHeader readHeadMeas(byte[] blobFile)
        {
            byte[] mHeader = blobFile;

            MemoryStream headerStream = new MemoryStream();
            headerStream.Write(mHeader, 0, mHeader.Length);
            headerStream.Seek(0, SeekOrigin.Begin);


            BinaryReader reader = new BinaryReader(headerStream);

            byte[] trash;

            spectrumHeader mSpectrumHeader = new spectrumHeader();
            mSpectrumHeader.Version = reader.ReadUInt16();

            trash = reader.ReadBytes(mByteBoundaryOffset);

            mSpectrumHeader.NumberOfLines = reader.ReadUInt16();
            mSpectrumHeader.NumberOfAverages = reader.ReadUInt16();
            mSpectrumHeader.Windowing = reader.ReadUInt32();
            mSpectrumHeader.MotorSpeed = reader.ReadDouble();
            mSpectrumHeader.Factor = reader.ReadDouble();
            mSpectrumHeader.Offset = reader.ReadDouble();
            mSpectrumHeader.HighFrequency = reader.ReadDouble();
            mSpectrumHeader.LowFrequency = reader.ReadDouble();
            mSpectrumHeader.Load = reader.ReadDouble();
            mSpectrumHeader.Detection = reader.ReadUInt32();
            mSpectrumHeader.Final = reader.ReadUInt32();
            mSpectrumHeader.Data = mHeader;

            return mSpectrumHeader;
        }
        private byte[] writeHeadMeas(spectrumHeader mSpectrumHeader)
        {

            MemoryStream headerWrite = new MemoryStream();

            byte[] trash = { 0, 0, 0, 0, 0, 0 };

            BinaryWriter writer = new BinaryWriter(headerWrite);
            writer.Write(mSpectrumHeader.Version);
            writer.Write(trash);
            writer.Write(mSpectrumHeader.NumberOfLines);
            writer.Write(mSpectrumHeader.NumberOfAverages);
            writer.Write(mSpectrumHeader.Windowing);
            writer.Write(mSpectrumHeader.MotorSpeed);
            writer.Write(mSpectrumHeader.Factor);
            writer.Write(mSpectrumHeader.Offset);
            writer.Write(mSpectrumHeader.HighFrequency);
            writer.Write(mSpectrumHeader.LowFrequency);
            writer.Write(mSpectrumHeader.Load);
            writer.Write(mSpectrumHeader.Detection);
            writer.Write(mSpectrumHeader.Final);

            byte[] write_result = headerWrite.ToArray();

            return write_result;
        }

        private double[] readReadingMeas(byte[] blobFile, spectrumHeader mSpectrumHeader)
        {
            byte[] mData = blobFile;
            double[] BinValues;
            short[] scaleValues = new short[mData.Length / sizeof(short)];
            Buffer.BlockCopy(mData, 0, scaleValues, 0, mData.Length);

            BinValues = new double[scaleValues.Length];

            for (int i = 0; i < scaleValues.Length; i++)
            {
                BinValues[i] = RestoreValue(scaleValues[i], mSpectrumHeader);
            }
            return BinValues;
        }
        private byte[] writeDataMeas(double[] BinValues, spectrumHeader mSpectrumHeader, bool _ConvertValue = true)
        {
            double[] BinValues_retornando = new double[BinValues.Length];
            byte[] retornando_para_byte = new byte[BinValues.Length * sizeof(short)];
            if (_ConvertValue)
            {
                for (int i = 0; i < BinValues.Length; i++)
                {
                    BinValues_retornando[i] = (short)ConvertValue(BinValues[i], mSpectrumHeader);
                }

                Buffer.BlockCopy(BinValues_retornando, 0, retornando_para_byte, 0, BinValues.Length * sizeof(short));
            }
            else
            {
                Buffer.BlockCopy(BinValues, 0, retornando_para_byte, 0, BinValues.Length * sizeof(short));
            }
            return retornando_para_byte;
        }

        private double RestoreValue(short iInput, spectrumHeader mSpectrumHeader)
        {
            return (double)(iInput * mSpectrumHeader.Factor);
        }
        private double ConvertValue(double iInput, spectrumHeader mSpectrumHeader)
        {
            return (double)(iInput / mSpectrumHeader.Factor);
        }

        private byte[] getByte(string Field, int MeasId, int Type = int.MinValue)
        {
            SqlConnection connection = new SqlConnection(Conex.ConnectionString);
            connection.Open();
            string sql = "SELECT " + Field + " FROM MEASREADING WHERE MEASID=" + MeasId + " AND READINGTYPE=" + Type;
            SqlCommand command = new SqlCommand(sql, connection);
            byte[] buffer = (byte[])command.ExecuteScalar();
            connection.Close();
            return buffer;

       }

    }
    public class spectrumHeader
    {
        public ushort Version {get;set; }
        public uint UnusedBuffer { get; set; }
        public ushort NumberOfLines { get; set; }
        public ushort NumberOfAverages { get; set; }
        public uint Windowing { get; set; }
        public double MotorSpeed { get; set; }
        public double Factor { get; set; }
        public double Offset { get; set; }
        public double HighFrequency { get; set; }
        public double LowFrequency { get; set; }
        public double Load { get; set; }
        public uint Detection { get; set; }
        public byte[] Data { get; set; }
        public uint Final { get; set; }
    }

    public static class Tech
    {
        public static string Techniques_String(Techniques Tec)
        {
            string _Return = null;
            System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
            switch (Tec)
            {
                case Techniques.Sensitive:
                    //if (ci.Name == "en-us")
                    //    _Return = "SENSITIVE";
                    //else
                    _Return = "SENSITIVA";
                    break;
                case Techniques.Vibration:
                    //if (ci.Name == "en-us")
                    //    _Return = "VIBRATION";
                    //else
                    _Return = "VIBRAÇÃO";
                    break;
                case Techniques.TrendOil:
                    //if (ci.Name == "en-us")
                    //    _Return = "OIL ANALYSIS";
                    //else
                    _Return = "ANÁLISE DE ÓLEO";
                    break;
                case Techniques.MCD:
                    _Return = "VIBRAÇÃO MCD";
                    break;
            }

            return _Return;
        }
        public static string Techniques_String_Short(Techniques Tec)
        {
            string _Return = null;

            switch (Tec)
            {
                case Techniques.Sensitive:
                    _Return = "S";
                    break;
                case Techniques.Vibration:
                    _Return = "V";
                    break;
                case Techniques.TrendOil:
                    _Return = "TO";
                    break;
                case Techniques.MCD:
                    _Return = "M";
                    break;
            }

            return _Return;
        }

        public static string[] DadTypes(AnalystConnection Connection, Techniques _tec)
        {
            string[] DadType = { };

            if (_tec == Techniques.Vibration)
            {
                DadType = new string[] { 
                                Registration.RegistrationId(Connection, "SKFCM_ASDD_MicrologDAD").ToString()
                            ,   Registration.RegistrationId(Connection, "SKFCM_ASDD_ImxDAD").ToString()
                            //,   "MCD"
                        };
            }

            if (_tec == Techniques.MCD)
            {
                DadType = new string[]
                {
                    "MCD"
                };
            }

            if (_tec == Techniques.Sensitive)
            {
                DadType = new string[] { 
                               Registration.RegistrationId(Connection, "selec * ").ToString()
                               , Registration.RegistrationId(Connection, "SKFCM_ASPT_MultipleInspection").ToString()
                        };
            }
            if (_tec == Techniques.Derivated)
            {
                DadType = new string[] { 
                                Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint").ToString()
                        };
            }
            if (_tec == Techniques.TrendOil)
            {
                DadType = new string[] { 
                                Registration.RegistrationId(Connection, "SKFCM_ASDD_OPEC_DAD").ToString()
                        };
            }
            if (_tec == Techniques.IMx)
            {
                DadType = new string[] { 
                               Registration.RegistrationId(Connection, "SKFCM_ASDD_ImxDAD").ToString()
                        };
            }
            return DadType;
        }
    }
    public static class Registration
    {
        static private AnalystConnection _Connection;
        static private bool _RegistrationLoaded = false;
        static private DataTable _RegistrationTable;
        //static private DataTable RegistrationTable { get { return (_RegistrationLoaded ? _RegistrationTable : null); } }

        static private bool _Load(AnalystConnection AnalystConnection, bool ForceReload = false)
        {
            if (ForceReload) _RegistrationLoaded = false;

            if (_Connection != AnalystConnection)
            {
                _RegistrationTable = new DataTable();
                _RegistrationLoaded = false;
                _Connection = AnalystConnection;
            }

            if (_Connection.IsConnected)
            {
                if (!_RegistrationLoaded)
                {
                    _RegistrationTable = AnalystConnection.DataTable("RegistrationId, Signature, Type, DefaultName, PluginId", "Registration");
                    if (_RegistrationTable.Rows.Count > 0)
                        _RegistrationLoaded = true;
                    else
                        _RegistrationTable = new DataTable();
                }
            }
            else
            {
                _RegistrationTable = new DataTable();
                _RegistrationLoaded = false;
            }

            return _RegistrationLoaded;
        }

        /// <summary>
        /// Returns the RegistrationId unique number for a given Signature string
        /// </summary>
        /// <param name="AnalystConnection">Analyst connection object</param>
        /// <param name="Signature">Signature string</param>
        /// <returns>RegistrationId from Registration table</returns>
        static public uint RegistrationId(AnalystConnection AnalystConnection, string Signature)
        {
            uint ReturnValue = 0;

            if (_Load(AnalystConnection, false))
            {
                DataRow[] RegistrationRows = _RegistrationTable.Select("Signature='" + Signature + "'");
                if (RegistrationRows.Length > 0)
                    ReturnValue = Convert.ToUInt32(RegistrationRows[0]["RegistrationId"]);
            }

            return ReturnValue;
        }

        /// <summary>
        /// Returns the Signature string for a given RegistrationId identifier
        /// </summary>
        /// <param name="AnalystConnection">Analyst connection object</param>
        /// <param name="RegistrationId">Registration unique number</param>
        /// <returns>Signature string from Registration table</returns>
        static public string Signature(AnalystConnection AnalystConnection, uint RegistrationId)
        {
            string ReturnValue = string.Empty;

            if (_Load(AnalystConnection, false))
            {
                DataRow[] RegistrationRows = _RegistrationTable.Select("RegistrationId=" + RegistrationId.ToString());
                if (RegistrationRows.Length > 0)
                    ReturnValue = RegistrationRows[0]["Signature"].ToString();
            }

            return ReturnValue;
        }

        //public static enum DADType
        //{
        //    MicrologInspector = "SKFCM_ASDD_MarlinDAD",
        //    Derivated = "SKFCM_ASDD_DerivedPoint",
        //    MicrologAnalyzer = "SKFCM_ASDD_MicrologDAD",
        //    IMx = "SKFCM_ASDD_ImxDAD"
        //}
    }

    /// <summary>
    /// Connection to SKF @ptitude Anayst database
    /// </summary>
    public class AnalystConnection : DBTools
    {
        private bool _IsLogged = false;
        public bool IsLogged { get { return _IsLogged; } }

        static private bool _UserTblLoaded = false;
        static private DataTable _UserTbl;

        private bool _UserTblRowLoaded = false;
        private DataRow _UserTblRow;

        public uint UserId { get { return (_IsLogged ? (uint)_UserTblRow["UserId"] : 0); } }
        public string UserName { get { return (_IsLogged ? _UserTblRow["LoginName"].ToString() : string.Empty); } }
        public uint SystemAccessDefId { get { return (_IsLogged ? (uint)_UserTblRow["SystemAccessDefId"] : 0); } }

        private string _Version;
        public string Version
        {
            get
            {
                if (IsConnected && string.IsNullOrEmpty(_Version))
                    _Version = SQLtoString("Version", "Tool", "Signature='SKFCM_MachineAnalystDbVersion'");

                return _Version;
            }
        }

        /// <summary>Returns true if user has "Administrator" or "Field Service" privilegies.</summary>
        public bool IsPowerUser { get { return (IsConnected && _IsLogged && ((SystemAccessDefId == 500) | (SystemAccessDefId == 400) | (SystemAccessDefId == 5) | (SystemAccessDefId == 4))); } }

        public string NoPassword
        {
            get
            {
                return
            char.ConvertFromUtf32(34) +
            char.ConvertFromUtf32(38) +
            char.ConvertFromUtf32(71) +
            char.ConvertFromUtf32(46) +
            char.ConvertFromUtf32(63) +
            char.ConvertFromUtf32(21) +
            char.ConvertFromUtf32(89) +
            char.ConvertFromUtf32(95) +
            char.ConvertFromUtf32(125) +
            char.ConvertFromUtf32(39) +
            char.ConvertFromUtf32(115) +
            char.ConvertFromUtf32(48);
            }
        }

        public AnalystConnection() { }
        public AnalystConnection(DBType DatabaseType, string DatabaseName, string DatabaseInitialCatalog, string DatabaseUserName, string DatabaseUserPassword)
        {
            _Connect(DatabaseType, DatabaseName, DatabaseInitialCatalog, DatabaseUserName, DatabaseUserPassword);
        }
        public AnalystConnection(DBType DatabaseType, string DatabaseName, string DatabaseInitialCatalog, string DatabaseUserName, string DatabaseUserPassword, string AnalystUserName, string AnalystUserPassword)
        {
            _Connect(DatabaseType, DatabaseName, DatabaseInitialCatalog, DatabaseUserName, DatabaseUserPassword);
            Login(AnalystUserName, AnalystUserPassword);
        }

        private bool _Connect(DBType DatabaseType, string DatabaseName, string DatabaseInitialCatalog, string DatabaseUserName, string DatabaseUserPassword)
        {
            DBType = DatabaseType;
            DBName = DatabaseName;
            InitialCatalog = DatabaseInitialCatalog;
            User = DatabaseUserName;
            Password = DatabaseUserPassword;

            return Connect();
        }

        public bool Detect()
        {
            GenericTools.DebugMsg("Detect(): Starting...");

            bool ReturnValue = false;
            GenericTools.LastError.Clear();

            if (!IsConnected)
            {
                try
                {
                    RegistryKey localKey;
                    RegistryKey localUser;
                    if (Environment.Is64BitOperatingSystem)
                    {
                        localKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                        localUser = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry64);
                    }

                    else
                    {
                        localKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
                        localUser = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry32);
                    }
                    string Analyst_DB_Config = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\5507EA8C684842143ACEB27C2373D26F").GetValue("C462F6AF40A63FA458C240DE02BC1C82").ToString();
                    string LastDataBase = localUser.OpenSubKey(@"S-1-5-21-2474387369-1653127123-1608106642-189664\Software\SKF Condition Monitoring\SKF Machine Analyst\Application Settings").GetValue("LastDbServer").ToString();
                    string LastDataBaseAccount = localUser.OpenSubKey(@"S-1-5-21-2474387369-1653127123-1608106642-189664\Software\SKF Condition Monitoring\SKF Machine Analyst\Application Settings").GetValue("LastDbAccount").ToString();

                    XmlDocument doc = new XmlDocument();
                    string File = Path.GetDirectoryName(Analyst_DB_Config) + "\\skfDbConnections.config";

                    doc.Load(File);
                    string db_connector_string = doc.InnerXml;
                    string[] db_connector_split1 = db_connector_string.Split('=');
                    string[] db_connector_split2 = db_connector_split1[4].Split('\"');
                    string db_connector = db_connector_split2[1];

                    if (db_connector == "System.Data.SqlClient")
                        this.DBType = DB.DBType.MSSQL;
                    else
                        this.DBType = DB.DBType.Oracle;

                    this.DBName = LastDataBase;
                    this.InitialCatalog = "SKFUser";
                    this.User = LastDataBaseAccount;
                    this.Password = "cm";

                    Connect();

                    if (IsConnected) ReturnValue = true;


                }
                catch (Exception ex)
                {
                    GenericTools.GetError("Detect(): " + ex.Message);
                }
                if (!ReturnValue) Close();
            }

            GenericTools.DebugMsg("Detect(): " + ReturnValue.ToString());

            return ReturnValue;
        }


        private bool _LoadUserTbl()
        {
            _UserTblLoaded = false;

            if (IsConnected)
                _UserTbl = DataTable("*", "UserTbl");

            _UserTblLoaded = (_UserTbl.Rows.Count > 0);

            return _UserTblLoaded;
        }

        private bool _LoadUserTblRow(uint AnalystUserId)
        {
            _UserTblRowLoaded = false;

            if (_UserTblLoaded)
            {
                DataRow[] UserTblRows = _UserTbl.Select("UserId=" + AnalystUserId.ToString());
                if (UserTblRows.Length > 0)
                {
                    _UserTblRow = UserTblRows[0];
                    _UserTblRowLoaded = true;
                }
            }

            if (!_UserTblRowLoaded) _UserTblRow = null;

            return _UserTblRowLoaded;
        }
        private bool _LoadUserTblRow(string AnalystUserName)
        {
            _UserTblRowLoaded = false;

            if (_UserTblLoaded)
            {
                DataRow[] UserTblRows = _UserTbl.Select("upper(LoginName)='" + AnalystUserName.ToString() + "'");
                if (UserTblRows.Length > 0)
                {
                    _UserTblRow = UserTblRows[0];
                    _UserTblRowLoaded = true;
                }
            }

            if (!_UserTblRowLoaded) _UserTblRow = null;

            return _UserTblRowLoaded;
        }

        /// <summary>Validate application user login.</summary>
        /// <param name="UserName">Application login name</param>
        /// <param name="Password">Application user password</param>
        /// <returns>True when user and password are validated</returns>
        public bool Login(string AnalystUserName, string AnalystPassword) { return Login(GetUserId(AnalystUserName), AnalystPassword); }
        /// <summary>Validate application user login.</summary>
        /// <param name="UserId">Application user id</param>
        /// <param name="Password">Application user password</param>
        /// <returns>True when user and password are validated</returns>
        public bool Login(uint AnalystUserId, string AnalystPassword)
        {
            GenericTools.DebugMsg("AppLogin(" + AnalystUserId.ToString() + ",*): Starting...");

            _IsLogged = false;

            try
            {
                if (!_UserTblRowLoaded) _LoadUserTblRow(AnalystUserId);

                if (_UserTblRowLoaded)
                    _IsLogged = (GenericTools.PassEncode(AnalystPassword) == _UserTblRow["Passwd"].ToString());
            }
            catch (Exception ex)
            {
                GenericTools.WriteLog("AppLogin(" + AnalystUserId.ToString() + ",*) error: " + ex.Message);
            }

            GenericTools.DebugMsg("AppLogin(" + AnalystUserId.ToString() + ",*): " + _IsLogged.ToString());
            return _IsLogged;
        }

        ///<summary>Get application user id.</summary>
        ///<param name="UserName">Application login name</param>
        ///<returns>Return value is lower than 1 if no user found.</returns>
        public uint GetUserId(string AnalystUserName)
        {
            GenericTools.DebugMsg("GetUserId(" + AnalystUserName + "): Starting...");

            uint ReturnValue = 0;

            try
            {
                if (!_UserTblLoaded) _LoadUserTbl();

                if (_UserTblLoaded)
                {
                    DataRow[] UserRows = _UserTbl.Select("upper(LoginName)='" + AnalystUserName.ToUpper() + "'");
                    if (UserRows.Length > 0)
                        ReturnValue = (uint)UserRows[0]["UserId"];
                }
            }
            catch (Exception ex)
            {
                GenericTools.WriteLog("GetUserId(" + AnalystUserName + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetUserId(" + AnalystUserName + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        /// <summary>Add notes to element</summary>
        /// <param name="OwnerId">Element unique id (TreeElemId)</param>
        /// <param name="Text">Notes text</param>
        /// <returns>Notes unique id</returns>


        /// <summary>Returns the list of hierarchies for a specific user.</summary>
        /// <param name="Sort">Sort list by name</param>
        /// <returns>Return value is a 2 columns table:
        /// <para>- TblSetId: Hierarchy id number</para>
        /// <para>- TblSetName: Hierarchy Name</para>
        /// </returns>
        public DataTable HierarchyList(bool Sort)
        {
            GenericTools.DebugMsg("HierarchyList(" + Sort.ToString() + "): Starting...");

            DataTable ReturnValue = new DataTable();

            try
            {
                if (IsConnected)
                    // if (IsPowerUser)
                    ReturnValue = DataTable("TblSetId, TblSetName", "TableSet" + (Sort ? " order by TblSetName" : ""));
                //  else
                //      ReturnValue = DataTable("TblSetId, TblSetName", "TableSet", "TblSetId in (select TblSetId from " + Owner + "HierAccess where UserId=" + UserId.ToString() + ")" + (Sort ? " order by TblSetName" : string.Empty));
            }
            catch (Exception ex)
            {
                GenericTools.WriteLog("HierarchyList(" + Sort.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("HierarchyList(" + Sort.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }

    }

    /// <summary>
    /// DEPRECATED Class
    /// Use AnalystConnection instead of this
    /// </summary>
    public class AnConnection : AnalystConnection
    {
        public static AnMeasurement[] MeasReadingElements = new AnMeasurement[0];
        public static DataTable RegistrationTable = new DataTable();
        private static bool bMeasReading_Loaded = false;
        public static string AppUserName = string.Empty;
        public static string AppPassword = string.Empty;
        public static Int32 AppUserId = -1;
        public static AnPoint[] Point = new AnPoint[0];

        /*  public string NoPassword =
            char.ConvertFromUtf32(34) +
            char.ConvertFromUtf32(38) +
            char.ConvertFromUtf32(71) +
            char.ConvertFromUtf32(46) +
            char.ConvertFromUtf32(63) +
            char.ConvertFromUtf32(21) +
            char.ConvertFromUtf32(89) +
            char.ConvertFromUtf32(95) +
            char.ConvertFromUtf32(125) +
            char.ConvertFromUtf32(39) +
            char.ConvertFromUtf32(115) +
            char.ConvertFromUtf32(48);
        */

        public AnConnection(string sAnConnection_ConnectionString)
        {
            Connect(sAnConnection_ConnectionString, DBType.MSSQL);
        }

        public AnConnection(string sAnConnection_ConnectionString, DBType dbType)
        {
            Connect(sAnConnection_ConnectionString, dbType);
            if (IsConnected)
            {
                ImportRegistration();
                InitializeMeasReading();
            }
        }
        public AnConnection(bool bAnConnection_Detect)
        {
            if (bAnConnection_Detect)
            {
                Detect();
                Connect();
                if (IsConnected)
                {
                    ImportRegistration();
                    InitializeMeasReading();
                }
            }
        }
        /// <summary>Create a new SKF @ptitude Analyst connection instance</summary>
        /// <param name="iDBType">Database type:
        /// <para>1 for Oracle</para>
        /// <para>2 for MSSQL Server</para>
        /// </param>
        /// <param name="sDBName">Database instance</param>
        /// <param name="sInitialCatalog">Initial catalog (MSSQL only)</param>
        /// <param name="sUser">Database user name</param>
        /// <param name="sPassword">Database user password</param>
        public AnConnection(DBType iDBType, string sDBName, string sInitialCatalog, string sUser, string sPassword)
        {
            if (iDBType == DBType.Oracle | iDBType == DBType.MSSQL) DBType = iDBType;
            if (sDBName != string.Empty) DBName = sDBName;
            if (sInitialCatalog != string.Empty) InitialCatalog = sInitialCatalog;
            if (sUser != string.Empty) User = sUser;
            if (sPassword != string.Empty) Password = sPassword;
            Connect();
            if (IsConnected)
            {
                ImportRegistration();
                InitializeMeasReading();
                GetDBVersion();
            }
        }
        public AnConnection() { }

        public bool Detect()
        {
            GenericTools.DebugMsg("Detect(): Starting...");

            bool ReturnValue = false;
            GenericTools.LastError.Clear();

            if (!IsConnected)
            {
                try
                {
                    RegistryKey oAnDB_ConnectionKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\SKF Condition Monitoring\\SKF Machine Analyst\\Application Settings");

                    string sAnDB_ConnectionString = string.Empty;
                    string sAnDB_ConnectionString_Aux = string.Empty;

                    int nPosition = 0;

                    if (oAnDB_ConnectionKey != null)
                    {
                        string KeyDbClassId = oAnDB_ConnectionKey.GetValue("DbClassId", string.Empty).ToString().ToUpper();
                        string KeyDbConnectName = oAnDB_ConnectionKey.GetValue("DbConnectName", string.Empty).ToString();
                        if ((string.IsNullOrEmpty(KeyDbClassId)) & (!string.IsNullOrEmpty(KeyDbConnectName))) KeyDbClassId = "{C1858FA3-8843-11D1-BB0D-0080C853CC11}";
                        switch (KeyDbClassId)
                        {
                            case "{74786FF5-A439-496B-BBC8-F93E73974301}": //SQL
                                DBType = DBType.MSSQL;
                                DBName = oAnDB_ConnectionKey.GetValue("DbConnectName").ToString();
                                InitialCatalog = string.Empty;
                                User = "SKFUser1";
                                Password = "cm";

                                sAnDB_ConnectionString = " " + oAnDB_ConnectionKey.GetValue("DbConnectFormat").ToString();
                                sAnDB_ConnectionString_Aux = sAnDB_ConnectionString.ToUpper();

                                nPosition = sAnDB_ConnectionString_Aux.IndexOf("INITIAL CATALOG=");
                                if (nPosition > 0)
                                {
                                    sAnDB_ConnectionString = sAnDB_ConnectionString.Substring(nPosition + 16);
                                    sAnDB_ConnectionString.Replace("\"", "'");
                                    if (sAnDB_ConnectionString.Substring(0, 1) == "'") sAnDB_ConnectionString = sAnDB_ConnectionString.Substring(1);
                                    if (sAnDB_ConnectionString.IndexOf("'") > 0) sAnDB_ConnectionString = sAnDB_ConnectionString.Substring(0, sAnDB_ConnectionString.IndexOf("'"));
                                    if (sAnDB_ConnectionString.IndexOf(";") > 0) sAnDB_ConnectionString = sAnDB_ConnectionString.Substring(0, sAnDB_ConnectionString.IndexOf(";"));
                                    InitialCatalog = sAnDB_ConnectionString;
                                    ReturnValue = true;
                                }
                                else
                                {
                                    DBType = 0;
                                    GenericTools.LastError.Message = "AnConnection.Detection(): InitialCatalog not detected";
                                    return ReturnValue;
                                }
                                break;

                            case "{C1858FA3-8843-11D1-BB0D-0080C853CC11}": //Oracle
                            default:
                                DBType = DBType.Oracle;
                                DBName = oAnDB_ConnectionKey.GetValue("DbConnectName").ToString();
                                InitialCatalog = string.Empty;
                                User = "SKFUser1";
                                Password = "cm";
                                ReturnValue = true;
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    GenericTools.GetError("Detect(): " + ex.Message);
                }
                if (!ReturnValue) Close();
            }

            GenericTools.DebugMsg("Detect(): " + ReturnValue.ToString());

            return ReturnValue;
        }

        private void InitializeMeasReading()
        {
            int MaxMeasId = SQLtoInt("max(MeasId)", "MeasReading");
            if (MaxMeasId > 0)
            {
                AnMeasurement[] Temp = new AnMeasurement[MaxMeasId];
                if (bMeasReading_Loaded)
                {
                    System.Array.Copy(MeasReadingElements, Temp, System.Math.Min(MeasReadingElements.Length, Temp.Length));
                }
                MeasReadingElements = Temp;
                bMeasReading_Loaded = true;
            }
        }

        //private static int[] aRegistration_RegistrationId;
        //private static string[] aRegistration_Signature;
        private static bool bRegistration_Scanned = false;
        private bool ImportRegistration()
        {
            if (!IsConnected) return false;
            try
            {
                GetDBVersion();
                RegistrationTable = RecordSet("select RegistrationId, Signature from " + Owner + "Registration");
                bRegistration_Scanned = true;
            }
            catch (Exception ex)
            {
                GenericTools.GetError("Error importing Registration table: " + ex.Message);
                return false;
            }
            return true;
            //DataTable oImportRegistration_DataTable = RecordSet("select RegistrationId, Signature from " + Owner + "Registration order by Signature Asc");
            //int[] aImportRegistration_RegistrationId = new int[oImportRegistration_DataTable.Rows.Count];
            //string[] aImportRegistration_Signature = new string[oImportRegistration_DataTable.Rows.Count];
            //for (int i = 0; i < oImportRegistration_DataTable.Rows.Count; i++)
            //{
            //aImportRegistration_RegistrationId[i] = Convert.ToInt16(oImportRegistration_DataTable.Rows[i][0]);
            //aImportRegistration_Signature[i] = oImportRegistration_DataTable.Rows[i][1].ToString();
            //}
            //oImportRegistration_DataTable.Dispose();

            // ADODB.Recordset oImportRegistration_RecordSet = new ADODB.Recordset();
            // oImportRegistration_RecordSet = RecordSet("select RegistrationId, Signature from " + Owner + "Registration order by Signature Asc");
            // if (!(oImportRegistration_RecordSet.State == 1)) return false;
            // int[] aImportRegistration_RegistrationId = new int[oImportRegistration_RecordSet.RecordCount];
            // string[] aImportRegistration_Signature = new string[oImportRegistration_RecordSet.RecordCount];
            // int nImportRegistration_Loop = 0;
            // 
            // oImportRegistration_RecordSet.MoveFirst();
            // while (!oImportRegistration_RecordSet.EOF)
            // {
            // aImportRegistration_RegistrationId[nImportRegistration_Loop] = Convert.ToInt16(oImportRegistration_RecordSet.Fields["RegistrationId"].Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            // aImportRegistration_Signature[nImportRegistration_Loop] = oImportRegistration_RecordSet.Fields["Signature"].Value.ToString();
            // 
            // nImportRegistration_Loop++;
            // oImportRegistration_RecordSet.MoveNext();
            // }
            // 
            // oImportRegistration_RecordSet.Close();
            // 
            //aRegistration_RegistrationId = aImportRegistration_RegistrationId;
            //aRegistration_Signature = aImportRegistration_Signature;
            //bRegistration_Scanned = true;
            // 
            //return true;
        }

        /// <summary>Get the registration id value for a provided signature string</summary>
        /// <param name="Signature">Signature string</param>
        /// <returns>Registration id value</returns>
        public Int32 GetRegistrationId(string Signature)
        {
            DateTime StartTime = DateTime.Now;
            int ReturnValue = 0;

            if (!bRegistration_Scanned)
            {
                bRegistration_Scanned = ImportRegistration();
                if (!bRegistration_Scanned)
                {
                    GenericTools.DebugMsg("GetRegistrationId(" + Signature + "): Error importing Registration table. - " + (DateTime.Now - StartTime).ToString());
                    return 0;
                }
            }
            try
            {
                DataRow[] RegistrationEntry = RegistrationTable.Select("Signature='" + Signature + "'");
                ReturnValue = Convert.ToInt32(RegistrationEntry[0]["RegistrationId"]);
            }
            catch (Exception ex)
            {
                GenericTools.GetError("GetRegistrationId(" + Signature + "): " + ex.Message);
                GenericTools.DebugMsg("GetRegistrationId(" + Signature + "): Error importing Registration table.");
            }
            GenericTools.DebugMsg("GetRegistrationId(" + Signature + "): " + ReturnValue.ToString() + " - " + (DateTime.Now - StartTime).ToString());
            return ReturnValue;
        }

        private DataTable TreeElem = new DataTable();
        private bool GetTreeElem_FirstRun = true;
        private bool GetTreeElem_IncludeDeleted = false;
        public DataTable GetTreeElem(bool IncludeDeleted, bool ForceReload) { return GetTreeElem(AppUserId, IncludeDeleted, ForceReload); }
        public DataTable GetTreeElem(Int32 UserId, bool IncludeDeleted, bool ForceReload)
        {
            GenericTools.DebugMsg("GetTreeElem(" + IncludeDeleted.ToString() + ", " + ForceReload.ToString() + "): Starting...");

            try
            {
                if (IsConnected)
                {
                    if (ForceReload | GetTreeElem_FirstRun | (IncludeDeleted != GetTreeElem_IncludeDeleted))
                    {
                        if (IncludeDeleted)
                        {
                            TreeElem = DataTable("TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger, Overdue, ChannelEnable", "TreeElem");
                        }
                        else
                        {
                            TreeElem = DataTable("TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger, Overdue, ChannelEnable", "TreeElem", "ParentId!=2147000000");
                        }
                    }

                    GetTreeElem_FirstRun = false;
                    GetTreeElem_IncludeDeleted = IncludeDeleted;
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetTreeElem(" + IncludeDeleted.ToString() + ", " + ForceReload.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetTreeElem(" + IncludeDeleted.ToString() + ", " + ForceReload.ToString() + ") returning table with " + TreeElem.Rows.Count.ToString() + " rows and " + TreeElem.Columns.Count.ToString() + " columns");
            return TreeElem;
        }

        private Hashtable TreeElemNodeKeys = new Hashtable();
        public Hashtable GetHashTreeElem(ref TreeView TreeViewElement, Int32 ElementId, Int32 SubLevels, bool ForceReload)
        {
            GenericTools.DebugMsg("GetHashTreeElem(" + ForceReload.ToString() + "): Starting...");
            TreeNode NewNode;
            TreeNode ParentNode;
            DataView SetListView;
            DataTable SetList;
            DataRow[] TreeElemSelect;
            DataTable HierList = GetUserHierarchyList(AppUserId, true);
            DataTable SubTree = new DataTable();

            if (ForceReload) TreeElemNodeKeys.Clear();

            TreeElemNodeKeys.Add("1", TreeViewElement.Nodes.Add("Hierarquias"));
            ParentNode = (TreeNode)TreeElemNodeKeys["1"];
            ParentNode.Name = "1";

            TreeElemNodeKeys.Add("2", TreeViewElement.Nodes.Add("Rotas"));
            ParentNode = (TreeNode)TreeElemNodeKeys["2"];
            ParentNode.Name = "2";

            TreeElemNodeKeys.Add("3", TreeViewElement.Nodes.Add("Espaços de trabalho"));
            ParentNode = (TreeNode)TreeElemNodeKeys["3"];
            ParentNode.Name = "3";

            try
            {
                GenericTools.DebugMsg("GetHashTreeElem(" + ForceReload.ToString() + "): Listing hierarchies...");
                if (ElementId < 1) // Raiz
                {
                    foreach (DataRow Row in HierList.Rows)
                    {
                        //}
                        //for (Int32 i = 0; i < HierList.Rows.Count; i++)
                        //{
                        // Entradas para grupo de hierarquia
                        ParentNode = (TreeNode)TreeElemNodeKeys["1"];
                        NewNode = ParentNode.Nodes.Add(Row["TblSetName"].ToString());
                        TreeElemSelect = TreeElem.Select("TblSetId=" + Row["TblSetId"].ToString() + " and ContainerType=1 and HierarchyType=1");
                        NewNode.Name = TreeElemSelect[0]["TreeElemId"].ToString();
                        TreeElemNodeKeys.Add("1." + TreeElemSelect[0]["TreeElemId"].ToString(), NewNode);
                        /*
                        if (SubLevels < 0)
                        {
                            SubTree = GetTree(GenericTools.StrToInt32(TreeElemSelect[0]["TreeElemId"].ToString()));
                            if (SubTree.Rows.Count > 0)
                            {
                                for (Int32 k = 0; k < SubTree.Rows.Count; k++)
                                {
                                    if (GenericTools.StrToInt32(SubTree.Rows[0]["ContainerType"].ToString()) < 4)
                                    {
                                        ParentNode = (TreeNode)TreeElemNodeKeys["1." + SubTree.Rows[k]["ParentId"].ToString()];
                                        NewNode = ParentNode.Nodes.Add(SubTree.Rows[k]["Name"].ToString());
                                        NewNode.Name = SubTree.Rows[k]["TreeElemId"].ToString();
                                        TreeElemNodeKeys.Add("1." + SubTree.Rows[k]["TreeElemId"].ToString(), NewNode);
                                    }
                                }
                            }
                        }
                        */

                        // Entradas para rotas
                        ParentNode = (TreeNode)TreeElemNodeKeys["2"];
                        NewNode = ParentNode.Nodes.Add(Row["TblSetName"].ToString());
                        NewNode.Name = "2." + Row["TblSetId"].ToString();
                        TreeElemNodeKeys.Add("2." + Row["TblSetId"].ToString(), NewNode);

                        try
                        {
                            ParentNode = (TreeNode)TreeElemNodeKeys["2." + Row["TblSetId"].ToString()];

                            SetListView = new DataView(TreeElem, "ContainerType=1 and HierarchyType=2 and TblSetId=" + Row["TblSetId"].ToString(), "Name", DataViewRowState.CurrentRows); //AnalystConnection.DataTable("TreeElemId, Name", "TreeElem", "ContainerType=1 and HierarchyType=2 and TblSetId=" + HierList.Rows[i]["TblSetId"].ToString() + " order by Name");
                            SetList = SetListView.ToTable();
                            for (Int32 j = 0; j < SetList.Rows.Count; j++)
                            {
                                NewNode = ParentNode.Nodes.Add(SetList.Rows[j]["Name"].ToString());
                                NewNode.Name = SetList.Rows[j]["TreeElemId"].ToString();
                                TreeElemNodeKeys.Add(SetList.Rows[j]["TreeElemId"].ToString(), NewNode);
                            }
                        }
                        catch (Exception ex)
                        {
                            GenericTools.DebugMsg("GetHashTreeElem(" + ForceReload.ToString() + ") error " + ex.Message);
                            TreeElemNodeKeys.Clear();
                        }

                        // Entradas para espaços de trabalho
                        ParentNode = (TreeNode)TreeElemNodeKeys["3"];
                        NewNode = ParentNode.Nodes.Add(Row["TblSetName"].ToString());
                        NewNode.Name = "3." + Row["TblSetId"].ToString();
                        TreeElemNodeKeys.Add("3." + Row["TblSetId"].ToString(), NewNode);

                        try
                        {
                            ParentNode = (TreeNode)TreeElemNodeKeys["3." + Row["TblSetId"].ToString()];

                            SetListView = new DataView(TreeElem, "ContainerType=1 and HierarchyType=3 and TblSetId=" + Row["TblSetId"].ToString(), "Name", DataViewRowState.CurrentRows);
                            SetList = SetListView.ToTable();
                            for (Int32 j = 0; j < SetList.Rows.Count; j++)
                            {
                                NewNode = ParentNode.Nodes.Add(SetList.Rows[j]["Name"].ToString());
                                NewNode.Name = SetList.Rows[j]["TreeElemId"].ToString();
                                TreeElemNodeKeys.Add(SetList.Rows[j]["TreeElemId"].ToString(), NewNode);
                            }
                        }
                        catch (Exception ex)
                        {
                            GenericTools.DebugMsg("GetHashTreeElem(" + ForceReload.ToString() + ") error " + ex.Message);
                            TreeElemNodeKeys.Clear();
                        }
                    }
                }
                else // Galho
                {

                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetHashTreeElem(" + ForceReload.ToString() + ") error " + ex.Message);
                TreeElemNodeKeys.Clear();
            }

            GenericTools.DebugMsg("GetHashTreeElem(" + ForceReload.ToString() + ") returning " + TreeElemNodeKeys.ToString());

            return TreeElemNodeKeys;
        }

        public string GetDBVersion()
        {
            DBVersion = SQLtoString("Version", "Tool", "Signature='SKFCM_MachineAnalystDbVersion'");
            GenericTools.DebugMsg("DBVersion: " + DBVersion);

            return DBVersion;
        }

        ///<summary>Returns the hierarchy name.</summary>
        ///<param name="TblSetId">Hierarchy id</param>
        public string GetHierarchyName(Int32 TblSetId)
        {
            GenericTools.DebugMsg("GetHierarchyName(" + TblSetId.ToString() + "): Starting");
            string ReturnValue = string.Empty;

            try
            {
                if (IsConnected) ReturnValue = SQLtoString("TblSetName", "TableSet", "TblSetId=" + TblSetId.ToString());
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetHierarchyName(" + TblSetId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetHierarchyName(" + TblSetId.ToString() + "): " + ReturnValue);
            return ReturnValue;
        }

        private static string GetHierarchyId_Last_HierarchyName = string.Empty;
        private static Int32 GetHierarchyId_Last_ResultValue = Int32.MinValue;
        ///<summary>Returns the hierarchy id based on its name.</summary>
        ///<param name="HierarchyName">Hierarchy name</param>
        public Int32 GetHierarchyId(string HierarchyName)
        {
            GenericTools.DebugMsg("GetHierarchyId(" + HierarchyName + "): Starting");
            Int32 ReturnValue = Int32.MinValue;

            if (HierarchyName == GetHierarchyId_Last_HierarchyName)
            {
                ReturnValue = GetHierarchyId_Last_ResultValue;
            }
            else
            {
                GetHierarchyId_Last_HierarchyName = HierarchyName;
                try
                {
                    if (IsConnected) ReturnValue = SQLtoInt("TblSetId", "TableSet", "upper(TblSetName)='" + HierarchyName.ToUpper() + "'");
                }
                catch (Exception ex)
                {
                    GenericTools.DebugMsg("GetHierarchyId(" + HierarchyName + ") error: " + ex.Message);
                }
            }
            GenericTools.DebugMsg("GetHierarchyId(" + HierarchyName + "): " + ReturnValue.ToString());
            GetHierarchyId_Last_ResultValue = ReturnValue;
            return ReturnValue;
        }

        ///<summary>Get application user id.</summary>
        ///<returns>Return value is lower than 1 if no user found.</returns>
        public Int32 GetAppUserId() { return GetAppUserId(AppUserName); }
        ///<summary>Get application user id.</summary>
        ///<param name="UserName">Application login name</param>
        ///<returns>Return value is lower than 1 if no user found.</returns>
        public Int32 GetAppUserId(string UserName)
        {
            GenericTools.DebugMsg("GetAppUserId(" + UserName + "): Starting");
            Int32 ReturnValue = Int32.MinValue;

            try
            {
                if (IsConnected & !string.IsNullOrEmpty(UserName)) ReturnValue = SQLtoInt("UserId", "UserTbl", "upper(LoginName)='" + UserName.ToUpper() + "'");
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetAppUserId(" + UserName + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetAppUserId(" + UserName + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        ///<summary>Get application user name.</summary>
        ///<returns>Returns empty if no user found.</returns>
        public string GetAppUserName() { return GetAppUserName(AppUserId); }
        ///<summary>Get application user name.</summary>
        ///<returns>Returns empty if no user found.</returns>
        ///<param name="UserId">Application user id</param>
        public string GetAppUserName(Int32 UserId)
        {
            GenericTools.DebugMsg("GetAppUserName(" + UserId.ToString() + "): Starting");
            string ReturnValue = string.Empty;

            try
            {
                if (IsConnected) ReturnValue = SQLtoString("LoginName", "UserTbl", "UserId=" + UserId.ToString());
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetAppUserName(" + UserId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetAppUserName(" + UserId.ToString() + "): " + ReturnValue);
            return ReturnValue;
        }

        /*
        ///<summary>Is power user (Administrator or Field Service).</summary>
        ///<returns>Returns true if user has "Administrator" or "Field Service" privilegies.</returns>
        public bool IsPowerUser() { return IsPowerUser(AppUserId); }
        ///<summary>Is power user (Administrator or Field Service).</summary>
        ///<param name="UserId">Application user id</param>
        ///<returns>Returns true if user has "Administrator" or "Field Service" privilegies.</returns>
        public bool IsPowerUser(Int32 UserId)
        {
            GenericTools.DebugMsg("IsPowerUser(" + UserId.ToString() + "): Starting");
            bool ReturnValue = false;

            try
            {
                if (IsConnected)
                {
                    int SystemAccessDefId = SQLtoInt("SystemAccessDefId", "UserTbl", "UserId=" + UserId.ToString());
                    ReturnValue = ((SystemAccessDefId == 500) | (SystemAccessDefId == 400) | (SystemAccessDefId == 5) | (SystemAccessDefId == 4));
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("IsPowerUser(" + UserId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("IsPowerUser(" + UserId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        */

        /// <summary>Returns the list of hierarchies for a specific user.</summary>
        /// <returns>Return value is a 2 columns table:
        /// <para>- TblSetId: Hierarchy id number</para>
        /// <para>- TblSetName: Hierarchy Name</para>
        /// </returns>
        public DataTable GetUserHierarchyList() { return GetUserHierarchyList(AppUserId, false); }
        /// <summary>Returns the list of hierarchies for a specific user.</summary>
        /// <param name="Sort">Sort list by name</param>
        /// <returns>Return value is a 2 columns table:
        /// <para>- TblSetId: Hierarchy id number</para>
        /// <para>- TblSetName: Hierarchy Name</para>
        /// </returns>
        public DataTable GetUserHierarchyList(bool Sort) { return GetUserHierarchyList(AppUserId, Sort); }
        /// <summary>Returns the list of hierarchies for a specific user.</summary>
        /// <param name="UserName">Application user login name</param>
        /// <returns>Return value is a 2 columns table:
        /// <para>- TblSetId: Hierarchy id number</para>
        /// <para>- TblSetName: Hierarchy Name</para>
        /// </returns>
        public DataTable GetUserHierarchyList(string UserName) { return GetUserHierarchyList(GetAppUserId(UserName), false); }
        /// <summary>Returns the list of hierarchies for a specific user.</summary>
        /// <param name="UserName">Application user login name</param>
        /// <param name="Sort">Sort list by name</param>
        /// <returns>Return value is a 2 columns table:
        /// <para>- TblSetId: Hierarchy id number</para>
        /// <para>- TblSetName: Hierarchy Name</para>
        /// </returns>
        public DataTable GetUserHierarchyList(string UserName, bool Sort) { return GetUserHierarchyList(GetAppUserId(UserName), Sort); }
        /// <summary>Returns the list of hierarchies for a specific user.</summary>
        /// <param name="UserId">Application user id</param>
        /// <returns>Return value is a 2 columns table:
        /// <para>- TblSetId: Hierarchy id number</para>
        /// <para>- TblSetName: Hierarchy Name</para>
        /// </returns>
        public DataTable GetUserHierarchyList(Int32 UserId) { return GetUserHierarchyList(UserId, false); }
        /// <summary>Returns the list of hierarchies for a specific user.</summary>
        /// <param name="UserId">Application user id</param>
        /// <param name="Sort">Sort list by name</param>
        /// <returns>Return value is a 2 columns table:
        /// <para>- TblSetId: Hierarchy id number</para>
        /// <para>- TblSetName: Hierarchy Name</para>
        /// </returns>
        public DataTable GetUserHierarchyList(Int32 UserId, bool Sort)
        {
            GenericTools.DebugMsg("GetUserHierarchyList(" + UserId.ToString() + "): Starting");
            DataTable ReturnValue = new DataTable();

            try
            {
                if (IsConnected)
                {
                    if (IsPowerUser)
                    {
                        ReturnValue = DataTable("TblSetId, TblSetName", "TableSet" + (Sort ? " order by TblSetName" : ""));
                    }
                    else
                    {
                        ReturnValue = DataTable("TblSetId, TblSetName", "TableSet", "TblSetId in (select TblSetId from " + Owner + "HierAccess where UserId=" + UserId.ToString() + ")" + (Sort ? " order by TblSetName" : string.Empty));
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetUserHierarchyList(" + UserId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetUserHierarchyList(" + UserId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        /// <summary>Validate application user login.</summary>
        /// <returns>True when user and password are validated</returns>
        public bool AppLogin() { return AppLogin(AppUserId, AppPassword); }
        /// <summary>Validate application user login.</summary>
        /// <param name="UserName">Application login name</param>
        /// <param name="Password">Application user password</param>
        /// <returns>True when user and password are validated</returns>
        public bool AppLogin(string UserName, string Password) { return AppLogin(GetAppUserId(UserName), Password); }
        /// <summary>Validate application user login.</summary>
        /// <param name="UserId">Application user id</param>
        /// <param name="Password">Application user password</param>
        /// <returns>True when user and password are validated</returns>
        public bool AppLogin(Int32 UserId, string Password)
        {
            GenericTools.DebugMsg("AppLogin(" + UserId.ToString() + "," + String.Empty.PadLeft(Password.Length, "*".ToCharArray()[0]) + "): Starting");
            bool ReturnValue = false;

            try
            {
                if (IsConnected)
                {
                    ReturnValue = (GenericTools.PassEncode(Password) == SQLtoString("Passwd", "UserTbl", "UserId=" + UserId.ToString()));
                    if (ReturnValue)
                    {
                        AppUserId = UserId;
                        AppPassword = Password;
                        AppUserName = GetUserLoginName(AppUserId);
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("AppLogin(" + UserId.ToString() + "," + String.Empty.PadLeft(Password.Length, "*".ToCharArray()[0]) + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("AppLogin(" + UserId.ToString() + "," + String.Empty.PadLeft(Password.Length, "*".ToCharArray()[0]) + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        /// <summary>Add new event on application event log</summary>
        /// <param name="Event">Event string</param>
        /// <returns>Returns event unique id</returns>
        public Int32 AppEventLog(string Event) { return AppEventLog(Event, 0, AppUserId); }
        /// <summary>Add new event on application event log</summary>
        /// <param name="Event">Event string</param>
        /// <param name="TableSetId">Hierarchy id</param>
        /// <returns>Returns event unique id</returns>
        public Int32 AppEventLog(string Event, Int32 TableSetId) { return AppEventLog(Event, TableSetId, AppUserId); }
        /// <summary>Add new event on application event log</summary>
        /// <param name="Event">Event string</param>
        /// <param name="TableSetId">Hierarchy id</param>
        /// <param name="UserId">Application user id</param>
        /// <returns>Returns event unique id</returns>
        public Int32 AppEventLog(string Event, Int32 TableSetId, Int32 UserId) { return AppEventLog(0, 0, 0, GetRegistrationId("SKFCM_ASDD_OPEC_DAD"), string.Empty, Event, GenericTools.DateTime(), TableSetId, GetRegistrationId("SKFCM_ASEG_DataCollection"), 0, UserId); }
        /// <summary>Add new event on application event log</summary>
        /// <param name="PointId">TreeElemId for point</param>
        /// <param name="HostId">Online system host id</param>
        /// <param name="DadType">Dad type</param>
        /// <param name="DeviceId">Device Id</param>
        /// <param name="Channel">Channel name</param>
        /// <param name="Event">Event string</param>
        /// <param name="EventDtg">Event date/time (YYYYMMDDHHMISS)</param>
        /// <param name="TableSetId">Hierarchy id</param>
        /// <param name="EventType">Event type</param>
        /// <param name="MeasId">Measurement id</param>
        /// <param name="UserId">Application user id</param>
        /// <returns>Returns event unique id</returns>
        public Int32 AppEventLog(Int32 PointId, Int32 HostId, Int32 DadType, Int32 DeviceId, string Channel, string Event, string EventDtg, Int32 TableSetId, Int32 EventType, Int32 MeasId, Int32 UserId)
        {
            GenericTools.DebugMsg("AppEventLog(\"" + Event + "\"): Starting");
            Int32 ReturnValue = Int32.MinValue;

            try
            {
                if (IsConnected)
                {
                    switch (DBType)
                    {
                        case DBType.Oracle: //Oracle
                            if (SQLExec(
                        "insert into " + Owner + "EventLog (EventId, PointId, HostId, DadType, DeviceId, Channel, Event, EventDtg, TableSetId, EventType, MeasId, UserId) Values (" +
                        "EventId_Seq.NextVal," + //EventId
                        PointId.ToString() + "," + //PointId
                        HostId.ToString() + "," + //HostId
                        DadType.ToString() + "," + //DadType
                        DeviceId.ToString() + "," + //DeviceId
                        (String.IsNullOrEmpty(Channel) ? "null," : "'" + Channel + "',") + //Channel
                        "'" + Event + "'," + //Event
                        "'" + EventDtg + "'," + //EventDtg
                        TableSetId.ToString() + "," + //TableSetId
                        EventType.ToString() + "," + //EventType
                        MeasId.ToString() + "," + //MeasId
                        UserId.ToString() + ")" //UserId
                        ))
                                ReturnValue = SQLtoInt("select " + Owner + "EventId_Seq.CurrVal from Dual");
                            break;

                        case DBType.MSSQL: //MSSQL
                            ReturnValue = SQLtoInt("insert into " + Owner + "EventLog (PointId, HostId, DadType, DeviceId, Channel, Event, EventDtg, TableSetId, EventType, MeasId, UserId) Values (" +
                        PointId.ToString() + "," + //PointId
                        HostId.ToString() + "," + //HostId
                        DadType.ToString() + "," + //DadType
                        DeviceId.ToString() + "," + //DeviceId
                        (String.IsNullOrEmpty(Channel) ? "null," : "'" + Channel + "',") + //Channel
                        "'" + Event + "'," + //Event
                        "'" + EventDtg + "'," + //EventDtg
                        TableSetId.ToString() + "," + //TableSetId
                        EventType.ToString() + "," + //EventType
                        MeasId.ToString() + "," + //MeasId
                        UserId.ToString() + ");select @@Identity;" //UserId
                        );
                            break;

                        default:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("AppEventLog(\"" + Event + "\") error: " + ex.Message);
            }

            GenericTools.DebugMsg("AppEventLog(\"" + Event + "\"): " + ReturnValue.ToString());
            return ReturnValue;
        }

        /// <summary>Returns application login name .</summary>
        public string GetUserLoginName() { return GetUserLoginName(AppUserId); }
        /// <summary>Returns application login name .</summary>
        /// <param name="UserId">Application user id</param>
        public string GetUserLoginName(Int32 UserId)
        {
            GenericTools.DebugMsg("GetUserLoginName(" + UserId.ToString() + "): Starting");
            string ReturnValue = String.Empty;

            try
            {
                if (IsConnected)
                {
                    ReturnValue = SQLtoString("LoginName", "UserTbl", "UserId=" + UserId.ToString());
                    if (!String.IsNullOrEmpty(ReturnValue) & UserId == AppUserId) AppUserName = ReturnValue;
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetUserLoginName(" + UserId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetUserLoginName(" + UserId.ToString() + "): " + ReturnValue);
            return ReturnValue;
        }

        /// <summary>Search machine on SKF @ptitude Analyst database</summary>
        /// <param name="TblSetId">Hierarchy id</param>
        /// <param name="GrandParentName">Grand Parent name</param>
        /// <param name="ParentName">Parent name</param>
        /// <param name="MachineName">Machine name</param>
        /// <returns>Table with TreeElemId</returns>
        public DataTable SearchMachine(Int32 TblSetId, string GrandParentName, string ParentName, string MachineName)
        {
            GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + GrandParentName + "\", \"" + ", \"" + ParentName + "\", \"" + MachineName + "\"): Starting");
            DataTable ReturnValue = new DataTable();

            try
            {
                if (IsConnected)
                {
                    ReturnValue = DataTable("TE1.TreeElemId", "TreeElem TE1", "TE1.HierarchyType=1 and TE1.ContainerType=3 and TE1.TblSetId=" + TblSetId.ToString() + " and exists (select TE2.TreeElemId from " + Owner + "TreeElemId TE2, TreeElemId TE3 where TE2.TreeElemId=TE1.ParentId and TE3.TreeElemId=TE2.ParentId and upper(TE2.Name)='" + ParentName.ToUpper() + "' and upper(TE3.Name)='" + GrandParentName.ToUpper() + "')");
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + GrandParentName + "\", \"" + ", \"" + ParentName + "\", \"" + MachineName + "\") error: " + ex.Message);
            }

            GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + GrandParentName + "\", \"" + ", \"" + ParentName + "\", \"" + MachineName + "\"): " + ReturnValue.ToString());
            return ReturnValue;
        }
        /// <summary>Search machine on SKF @ptitude Analyst database</summary>
        /// <param name="TblSetId">Hierarchy id</param>
        /// <param name="ParentName">Parent name</param>
        /// <param name="MachineName">Machine name</param>
        /// <returns>Table with TreeElemId</returns>
        public DataTable SearchMachine(Int32 TblSetId, string ParentName, string MachineName)
        {
            GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + ParentName + "\", \"" + MachineName + "\"): Starting");
            DataTable ReturnValue = new DataTable();

            try
            {
                if (IsConnected)
                {
                    ReturnValue = DataTable("TE1.TreeElemId", "TreeElem TE1", "upper(TE1.NAME) = '" + MachineName.ToUpper() + "' and TE1.HierarchyType=1 and TE1.ContainerType=3 and TE1.TblSetId=" + TblSetId.ToString() + " and exists (select TE2.TreeElemId from " + Owner + "TreeElem TE2 where TE2.TreeElemId=TE1.ParentId and upper(TE2.Name)='" + ParentName.ToUpper() + "')");
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + ParentName + "\", \"" + MachineName + "\") error: " + ex.Message);
            }

            GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + ParentName + "\", \"" + MachineName + "\"): " + ReturnValue.ToString());
            return ReturnValue;
        }
        /// <summary>Search machine on SKF @ptitude Analyst database</summary>
        /// <param name="TblSetId">Hierarchy id</param>
        /// <param name="ParentId">Parent id</param>
        /// <param name="MachineName">Machine name</param>
        /// <returns>Table with TreeElemId</returns>
        public DataTable SearchMachine(Int32 TblSetId, Int32 ParentId, string MachineName)
        {
            GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + ParentId.ToString() + "\", \"" + MachineName + "\"): Starting");
            DataTable ReturnValue = new DataTable();

            try
            {
                if (IsConnected)
                {
                    ReturnValue = DataTable("TreeElemId", "TreeElem", "HierarchyType=1 and ContainerType=3 and TblSetId=" + TblSetId.ToString() + " and ParentId=" + ParentId.ToString());
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + ParentId.ToString() + "\", \"" + MachineName + "\") error: " + ex.Message);
            }

            GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + ParentId.ToString() + "\", \"" + MachineName + "\"): " + ReturnValue.ToString());
            return ReturnValue;
        }
        /// <summary>Search machine on SKF @ptitude Analyst database</summary>
        /// <param name="TblSetId">Hierarchy id</param>
        /// <param name="MachineName">Machine name</param>
        /// <returns>Table with TreeElemId</returns>
        public DataTable SearchMachine(Int32 TblSetId, string MachineName)
        {
            GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + MachineName + "\"): Starting");
            DataTable ReturnValue = new DataTable();

            try
            {
                if (IsConnected)
                {
                    ReturnValue = DataTable("TreeElemId", "TreeElem", "HierarchyType=1 and ContainerType=3 and TblSetId=" + TblSetId.ToString());
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + MachineName + "\") error: " + ex.Message);
            }

            GenericTools.DebugMsg("SearchMachine(" + TblSetId.ToString() + ", \"" + MachineName + "\"): " + ReturnValue.ToString());
            return ReturnValue;
        }

        /// <summary>Search point into a specific machine</summary>
        /// <param name="ParentId">Machine id</param>
        /// <param name="PointName">Point name</param>
        /// <returns>Returns point unique id (TreeElemId)</returns>
        public Int32 SearchPoint(Int32 ParentId, string PointName)
        {
            GenericTools.DebugMsg("SearchPoint(" + ParentId.ToString() + ", \"" + PointName + "\"): Starting");
            Int32 ReturnValue = Int32.MinValue;

            try
            {
                if (IsConnected)
                {
                    ReturnValue = SQLtoInt("TreeElemId", "TreeElem", "HierarchyType=1 and ContainerType=4 and ParentId=" + ParentId.ToString() + " and upper(Name)='" + PointName.ToUpper() + "'");
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("SearchPoint(" + ParentId.ToString() + ", \"" + PointName + "\") error: " + ex.Message);
            }

            GenericTools.DebugMsg("SearchPoint(" + ParentId.ToString() + ", \"" + PointName + "\"): " + ReturnValue.ToString());
            return ReturnValue;
        }

        /// <summary>Store a overall value on SKF @ptitude Analyst database</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <param name="OverallValue">Scalar value do be stored</param>
        /// <returns>Measurement unique id (MeasId)
        /// <para>Returns negative id if measurement already exists on database</para>
        /// </returns>
        public Int32 StoreOverallData(Int32 PointId, float OverallValue) { return StoreOverallData(PointId, DateTime.Now, OverallValue); }
        /// <summary>Store a overall value on SKF @ptitude Analyst database</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <param name="TimeStamp">Measurement time stamp</param>
        /// <param name="OverallValue">Scalar value do be stored</param>
        /// <returns>Measurement unique id (MeasId)
        /// <para>Returns negative id if measurement already exists on database</para>
        /// </returns>
        public Int32 StoreOverallData(Int32 PointId, DateTime TimeStamp, float OverallValue) { return StoreOverallData(PointId, GenericTools.DateTime(TimeStamp), OverallValue); }
        /// <summary>Store a overall value on SKF @ptitude Analyst database</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <param name="DataDtg">Measurement time stamp (YYYYMMDDHHMISS)</param>
        /// <param name="OverallValue">Scalar value do be stored</param>
        /// <returns>Measurement unique id (MeasId)
        /// <para>Returns negative id if measurement already exists on database</para>
        /// </returns>
        public Int32 StoreOverallData(Int32 PointId, string DataDtg, float OverallValue)
        {
            GenericTools.DebugMsg("StoreOverallData(" + PointId.ToString() + ", \"" + DataDtg + "\", " + OverallValue.ToString() + "): Starting");
            Int32 ReturnValue = 0;

            try
            {
                if (IsConnected)
                {
                    Int32 LastMeasId = SQLtoInt("min(MeasId)", "MeasDtsRead", "PointId=" + PointId.ToString() + " and DataDtg='" + DataDtg + "' and ReadingType=" + GetRegistrationId("SKFCM_ASMD_Overall").ToString() + " and round(OverallValue,4)=round(" + OverallValue.ToString().Replace(",", ".") + ",4)");
                    if (LastMeasId > 0)
                    {
                        ReturnValue = LastMeasId * (-1);
                    }
                    else
                    {
                        switch (DBType)
                        {
                            case DBType.Oracle: //Oracle
                                if (SQLExec("insert into " + Owner + "Measurement (MEASID, POINTID, ALARMLEVEL, STATUS, INCLUDE, OPERNAME, SERIALNO, DATADTG, MEASUREMENTTYPE) values (" +
                                    "MeasId_Seq.NextVal," + //MEASID, 
                                    PointId.ToString() + "," + //POINTID, 
                                    "null," + //ALARMLEVEL, 
                                    "1," + //STATUS, 
                                    "1," + //INCLUDE, 
                                    "'" + GetAppUserName(AppUserId) + "'," +//OPERNAME, 
                                    "null," + //SERIALNO, 
                                    "'" + DataDtg + "'," + //DATADTG, 
                                    "0" + //MEASUREMENTTYPE
                                    ")"))
                                {
                                    ReturnValue = SQLtoInt("select " + Owner + "MeasId_Seq.CurrVal from Dual");
                                    if (SQLExec("insert into " + Owner + "MeasReading (READINGID, MEASID, POINTID, READINGTYPE, CHANNEL, OVERALLVALUE, EXDWORDVAL1, EXDOUBLEVAL1, EXDOUBLEVAL2, EXDOUBLEVAL3, READINGHEADER, READINGDATA) values (" +
                                        "MeasRdngId_Seq.NextVal," + //READINGID
                                        ReturnValue.ToString() + "," + //MEASID
                                        PointId.ToString() + "," + //POINTID
                                        GetRegistrationId("SKFCM_ASMD_Overall").ToString() + "," + //READINGTYPE
                                        "1," + //CHANNEL
                                        OverallValue.ToString().Replace(",", ".") + "," + //OVERALLVALUE
                                        "0," + //EXDWORDVAL1
                                        "0," + //EXDOUBLEVAL1
                                        "0," + //EXDOUBLEVAL2
                                        "0," + //EXDOUBLEVAL3
                                        "null," + //READINGHEADER
                                        "null" + //READINGDATA
                                        ")"))
                                    {
                                        Int32 ReadingId = SQLtoInt("select " + Owner + "MeasRdngId_Seq.CurrVal from Dual");
                                        SQLExec("delete from " + Owner + "MeasAlarm where PointId=" + PointId.ToString() + " and AlarmType=" + GetRegistrationId("SKFCM_ASAT_Overall").ToString() + " and TableName='ScalarAlarm'");
                                        if (!SQLExec("insert into " + Owner + "MeasAlarm (MeasAlarmId, PointId, ReadingId, AlarmId, AlarmType, AlarmLevel, ResultType, TableName, MeasId, Channel) values (" +
                                            "MeasAlarmId_Seq.NextVal," + //MeasAlarmId
                                            PointId.ToString() + "," + //PointId
                                            ReadingId.ToString() + "," + //ReadingId
                                            GetScalarAlarmId(PointId).ToString() + "," + //AlarmId
                                            GetRegistrationId("SKFCM_ASAT_Overall").ToString() + "," + //AlarmType
                                            GetScalarAlarmLevel(PointId, OverallValue)[0].ToString() + "," + //AlarmLevel
                                            GetScalarAlarmLevel(PointId, OverallValue)[1].ToString() + "," + //ResultType
                                            "'ScalarAlarm'," + //TableName
                                            ReturnValue.ToString() + "," + //MeasId
                                            "1" + //Channel
                                            ")"))
                                        {
                                            SQLExec("delete from " + Owner + "MeasReading where ReadingId=" + ReadingId.ToString());
                                            SQLExec("delete from " + Owner + "Measurement where MeasId=" + ReturnValue.ToString());
                                            ReturnValue = Int32.MinValue;
                                        }
                                    }
                                    else
                                    {
                                        SQLExec("delete from " + Owner + "Measurement where MeasId=" + ReturnValue.ToString());
                                        ReturnValue = Int32.MinValue;
                                    }
                                }
                                else
                                {
                                    ReturnValue = Int32.MinValue;
                                }
                                break;

                            case DBType.MSSQL: // MSSQL
                                ReturnValue = SQLtoInt("insert into " + Owner + "Measurement (POINTID, ALARMLEVEL, STATUS, INCLUDE, OPERNAME, SERIALNO, DATADTG, MEASUREMENTTYPE) values (" +
                                    PointId.ToString() + "," + //POINTID, 
                                    "null," + //ALARMLEVEL, 
                                    "1," + //STATUS, 
                                    "1," + //INCLUDE, 
                                    "'" + GetAppUserName(AppUserId) + "'," +//OPERNAME, 
                                    "null," + //SERIALNO, 
                                    "'" + DataDtg + "'," + //DATADTG, 
                                    "0" + //MEASUREMENTTYPE
                                    ");select @@Identity;");
                                if (ReturnValue > 0)
                                {
                                    Int32 ReadingId = SQLtoInt("insert into " + Owner + "MeasReading (MEASID, POINTID, READINGTYPE, CHANNEL, OVERALLVALUE, EXDWORDVAL1, EXDOUBLEVAL1, EXDOUBLEVAL2, EXDOUBLEVAL3, READINGHEADER, READINGDATA) values (" +
                                        ReturnValue.ToString() + "," + //MEASID
                                        PointId.ToString() + "," + //POINTID
                                        GetRegistrationId("SKFCM_ASMD_Overall").ToString() + "," + //READINGTYPE
                                        "1," + //CHANNEL
                                        OverallValue.ToString().Replace(",", ".") + "," + //OVERALLVALUE
                                        "0," + //EXDWORDVAL1
                                        "0," + //EXDOUBLEVAL1
                                        "0," + //EXDOUBLEVAL2
                                        "0," + //EXDOUBLEVAL3
                                        "null," + //READINGHEADER
                                        "null" + //READINGDATA
                                        ");select @@Identity;");
                                    if (ReadingId > 0)
                                    {
                                        SQLExec("delete from " + Owner + "MeasAlarm where PointId=" + PointId.ToString() + " and AlarmType=" + GetRegistrationId("SKFCM_ASAT_Overall").ToString() + " and TableName='ScalarAlarm'");
                                        if (SQLtoInt("insert into " + Owner + "MeasAlarm (PointId, ReadingId, AlarmId, AlarmType, AlarmLevel, ResultType, TableName, MeasId, Channel) values (" +
                                            PointId.ToString() + "," + //PointId
                                            ReadingId.ToString() + "," + //ReadingId
                                            GetScalarAlarmId(PointId).ToString() + "," + //AlarmId
                                            GetRegistrationId("SKFCM_ASAT_Overall").ToString() + "," + //AlarmType
                                            GetScalarAlarmLevel(PointId, OverallValue)[0].ToString() + "," + //AlarmLevel
                                            GetScalarAlarmLevel(PointId, OverallValue)[1].ToString() + "," + //ResultType
                                            "'ScalarAlarm'," + //TableName
                                            ReturnValue.ToString() + "," + //MeasId
                                            "1" + //Channel
                                            ");select @@Identity;") < 1)
                                        {
                                            SQLExec("delete from " + Owner + "MeasReading where ReadingId=" + ReadingId.ToString());
                                            SQLExec("delete from " + Owner + "Measurement where MeasId=" + ReturnValue.ToString());
                                            ReturnValue = Int32.MinValue;
                                        }
                                    }
                                    else
                                    {
                                        SQLExec("delete from " + Owner + "Measurement where MeasId=" + ReturnValue.ToString());
                                        ReturnValue = Int32.MinValue;
                                    }
                                }
                                else
                                {
                                    ReturnValue = Int32.MinValue;
                                }
                                break;

                            default: // Not implemented
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("StoreOverallData(" + PointId.ToString() + ", \"" + DataDtg + "\", " + OverallValue.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("StoreOverallData(" + PointId.ToString() + ", \"" + DataDtg + "\", " + OverallValue.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        /// <summary>Get scalar alarm id for a given point</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <returns>Alarm unique id for ScalarAlarmId table</returns>
        public Int32 GetScalarAlarmId(Int32 PointId)
        {
            GenericTools.DebugMsg("GetScalarAlarmId(" + PointId.ToString() + "): Starting");
            Int32 ReturnValue = Int32.MinValue;

            try
            {
                if (IsConnected)
                {
                    ReturnValue = SQLtoInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString() + " and Type=" + GetRegistrationId("SKFCM_ASAT_Overall").ToString());
                    if (ReturnValue < 1) ReturnValue = SQLtoInt("ScalarAlrmId", "ScalarAlarm", "ElementId=" + PointId.ToString());
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetScalarAlarmId(" + PointId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetScalarAlarmId(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        /// <summary>Get inspection alarm id for a given point</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <returns>Alarm unique id for InspectionAlarm table</returns>
        public Int32 GetInspectionAlarmId(Int32 PointId)
        {
            GenericTools.DebugMsg("GetInspectionAlarmId(" + PointId.ToString() + "): Starting");
            Int32 ReturnValue = Int32.MinValue;

            try
            {
                if (IsConnected)
                {
                    ReturnValue = SQLtoInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString() + " and Type=" + GetRegistrationId("SKFCM_ASAT_Inspection").ToString());
                    if (ReturnValue < 1) ReturnValue = SQLtoInt("InspectionAlrmId", "GetInspectionAlarm", "ElementId=" + PointId.ToString());
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetInspectionAlarmId(" + PointId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetInspectionAlarmId(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        public int GetInspectionResultAlarmFlag(Int32 PointId, int tInspecionResult)
        {
            GenericTools.DebugMsg("GetInspectionResultAlarmFlag(): Starting...");

            int ReturnValue = 0;

            Int32 InspectionAlrmId = SQLtoInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString());
            if (InspectionAlrmId < 1) InspectionAlrmId = SQLtoInt("InspectionAlrmId", "InspectionAlarm", "ElementId=" + PointId.ToString());

            for (int i = 0; i < 5; i++)
                if ((tInspecionResult & (int)Math.Pow(2, i)) > 0) ReturnValue = Math.Max(ReturnValue, SQLtoInt("AlarmLevel" + (i + 1).ToString(), "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString()));

            GenericTools.DebugMsg("GetInspectionResultAlarmFlag(): " + ReturnValue);

            return ReturnValue;
        }

        /// <summary>Get MCD alarm id for a given point</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <returns>Alarm unique id for MCDAlarm table</returns>
        public Int32 GetMCDAlarmId(Int32 PointId)
        {
            GenericTools.DebugMsg("GetMCDAlarmId(" + PointId.ToString() + "): Starting");
            Int32 ReturnValue = Int32.MinValue;

            try
            {
                if (IsConnected)
                {
                    ReturnValue = SQLtoInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString() + " and Type=" + GetRegistrationId("SKFCM_ASAT_MCD").ToString());
                    if (ReturnValue < 1) ReturnValue = SQLtoInt("MCDAlarmId", "MCDAlarm", "ElementId=" + PointId.ToString());
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetMCDAlarmId(" + PointId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetMCDAlarmId(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        public int GetMCDResultAlarmFlag(Int32 PointId, int AlarmType, Int32 ReadingId)
        {
            GenericTools.DebugMsg("GetMCDResultAlarmFlag(): Starting...");

            int ReturnValue = 2;

            float MeasuredValue = MeasuredValue = SQLtoFloat("ExDoubleVal" + AlarmType.ToString(), "MeasReading", "ReadingId=" + ReadingId.ToString());

            Int32 MCDAlarmId = GetMCDAlarmId(PointId);

            string AlarmPrefix = string.Empty;

            switch (AlarmType)
            {
                case 1:
                    AlarmPrefix = "TEMPERATURE";
                    break;

                case 2:
                    AlarmPrefix = "VELOCITY";
                    break;

                case 3:
                    AlarmPrefix = "ENVACC";
                    break;
            }

            float AlertHi = SQLtoFloat(AlarmPrefix + "ALERTHI", "MCDAlarm", "MCDAlarmId=" + MCDAlarmId.ToString());
            float DangerHi = SQLtoFloat(AlarmPrefix + "DANGERHI", "MCDAlarm", "MCDAlarmId=" + MCDAlarmId.ToString());

            if (MeasuredValue >= AlertHi) ReturnValue = 3;
            if (MeasuredValue >= DangerHi) ReturnValue = 4;

            GenericTools.DebugMsg("GetMCDResultAlarmFlag(): " + ReturnValue);

            return ReturnValue;
        }


        /// <summary>Get alarm level for specific value in a given point</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <returns>Alarm level:
        /// <para>0: No alarm info</para>
        /// <para>1: Out of alarm</para>
        /// <para>2: Alert</para>
        /// <para>3: Danger</para>
        /// </returns>
        public int[] GetScalarAlarmLevel(Int32 PointId) { return GetScalarAlarmLevel(PointId, SQLtoFloat("OverallValue", "MeasDtsRead", "PointId=" + PointId.ToString() + " and ReadingType=" + GetRegistrationId("SKFCM_ASMD_Overall").ToString() + " order by DataDtg desc")); }
        //public int GetScalarAlarmLevel(Int32 PointId, Int32 MeasId) { return GetScalarAlarmLevel(PointId, SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + GetRegistrationId("SKFCM_ASMD_Overall")));}
        /// <summary>Get alarm level for specific value in a given point</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <param name="ReadingId">Reading unique id</param>
        /// <returns>Alarm level:
        /// <para>0: No alarm info</para>
        /// <para>1: Out of alarm</para>
        /// <para>2: Alert</para>
        /// <para>3: Danger</para></returns>
        public int[] GetScalarAlarmLevel(Int32 PointId, Int32 ReadingId) { return GetScalarAlarmLevel(PointId, SQLtoFloat("OverallValue", "MeasReading", "ReadingId=" + ReadingId.ToString())); }
        /// <summary>Get alarm level for specific value in a given point</summary>
        /// <param name="PointId">Point unique id (TreeElemId)</param>
        /// <param name="OverallValue">Overall value measured</param>
        /// <returns>Alarm level:
        /// <para>0: No alarm info</para>
        /// <para>1: Out of alarm</para>
        /// <para>2: Alert</para>
        /// <para>3: Danger</para></returns>
        public int[] GetScalarAlarmLevel(Int32 PointId, float OverallValue)
        {
            GenericTools.DebugMsg("GetScalarAlarmLevel(" + PointId.ToString() + ", " + OverallValue.ToString() + "): Starting");
            int[] ReturnValue = new int[2] { 2, 0 };

            try
            {
                if (IsConnected)
                {
                    bool EnableDangerHi = false;
                    bool EnableDangerLo = false;
                    bool EnableAlertHi = false;
                    bool EnableAlertLo = false;
                    int AlarmMethod = 0;
                    float DangerLo = 0;
                    float DangerHi = 0;
                    float AlertLo = 0;
                    float AlertHi = 0;

                    DataTable ScalarAlarm = DataTable("ScalarAlrmId, ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi", "ScalarAlarm", "ScalarAlrmId=" + GetScalarAlarmId(PointId).ToString());
                    if (ScalarAlarm.Rows.Count > 0)
                    {
                        EnableDangerHi = (Convert.ToInt16(ScalarAlarm.Rows[0]["EnableDangerHi"]) == 1);
                        EnableDangerLo = (Convert.ToInt16(ScalarAlarm.Rows[0]["EnableDangerLo"]) == 1);
                        EnableAlertHi = (Convert.ToInt16(ScalarAlarm.Rows[0]["EnableAlertHi"]) == 1);
                        EnableAlertLo = (Convert.ToInt16(ScalarAlarm.Rows[0]["EnableAlertLo"]) == 1);
                        AlarmMethod = Convert.ToInt16(ScalarAlarm.Rows[0]["AlarmMethod"]);
                        DangerLo = (float)Convert.ToDouble(ScalarAlarm.Rows[0]["DangerLo"]);
                        DangerHi = (float)Convert.ToDouble(ScalarAlarm.Rows[0]["DangerHi"]);
                        AlertLo = (float)Convert.ToDouble(ScalarAlarm.Rows[0]["AlertLo"]);
                        AlertHi = (float)Convert.ToDouble(ScalarAlarm.Rows[0]["AlertHi"]);

                        switch (AlarmMethod)
                        {
                            case 1: //Level
                                if (EnableAlertHi & (OverallValue >= AlertHi)) ReturnValue = new int[2] { 3, 1 };
                                if (EnableDangerHi & (OverallValue >= DangerHi)) ReturnValue = new int[2] { 4, 3 };
                                break;

                            case 2: //Window
                                if (EnableDangerLo & (OverallValue >= DangerLo)) ReturnValue = new int[2] { 4, 4 };
                                if (EnableDangerHi & (OverallValue <= DangerHi)) ReturnValue = new int[2] { 4, 3 };
                                if (EnableAlertLo & (OverallValue >= AlertLo)) ReturnValue = new int[2] { 3, 2 };
                                if (EnableAlertHi & (OverallValue <= AlertHi)) ReturnValue = new int[2] { 3, 1 };
                                break;

                            case 3: //Out of window
                                if (EnableAlertHi & (OverallValue >= AlertHi)) ReturnValue = new int[2] { 3, 1 };
                                if (EnableAlertLo & (OverallValue <= AlertLo)) ReturnValue = new int[2] { 3, 2 };
                                if (EnableDangerHi & (OverallValue >= DangerHi)) ReturnValue = new int[2] { 4, 3 };
                                if (EnableDangerLo & (OverallValue <= DangerLo)) ReturnValue = new int[2] { 4, 4 };
                                break;

                            default: //None
                                ReturnValue = new int[2] { 2, 0 };
                                break;
                        }
                    }
                    /*
                    else
                    {
                        ReturnValue = new int[2];
                        if (GetScalarStdAlarmLevel(PointId, ref EnableDangerHi, ref EnableDangerLo, ref EnableAlertHi, ref EnableAlertLo, ref AlarmMethod, ref DangerLo, ref DangerHi, ref AlertLo, ref AlertHi))
                        {
                            GenericTools.DebugMsg("GetScalarAlarmLevel(" + PointId.ToString() + ", " + OverallValue.ToString() + "): Can't find alarm settings");
                        }
                    }
                    */
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetScalarAlarmLevel(" + PointId.ToString() + ", " + OverallValue.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetScalarAlarmLevel(" + PointId.ToString() + ", " + OverallValue.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        // ************ ATENCAO *************
        // Revisar código desta funcao GetScalarStdAlarmLevel. Atual código ajustado para emergencia Petrobras
        private bool GetScalarStdAlarmLevel(Int32 PointId, ref bool EnableDangerHi, ref bool EnableDangerLo, ref bool EnableAlertHi, ref bool EnableAlertLo, ref int AlarmMethod, ref float DangerLo, ref float DangerHi, ref float AlertLo, ref float AlertHi)
        {
            bool ReturnValue = false;
            GenericTools.DebugMsg("GetScalarStdAlarmLevel(" + PointId.ToString() + "): Starting...");
            //É ponto manual?
            if (SQLtoInt("count(*)", "Point", "ElementId=" + PointId.ToString() + " and FieldId=" + GetRegistrationId("SKFCM_ASPF_Dad_Id") + " and ValueString='" + GetRegistrationId("SKFCM_ASDD_ManualDAD") + "'") > 0)
            {
                switch (SQLtoString("Signature", "Registration", "RegistrationId=" + SQLtoString("ValueString", "Point", "ElementId=" + PointId.ToString() + " and FieldId=" + GetRegistrationId("SKFCM_ASPF_Point_Type_Id"))))
                {
                    case "SKFCM_ASPT_Wildcard":
                        AlarmMethod = 1;
                        EnableDangerHi = true;
                        DangerHi = (float)(3.5);
                        EnableAlertHi = true;
                        AlertHi = (float)(2);
                        break;

                    case "SKFCM_ASPT_Temperature":
                        AlarmMethod = 1;
                        EnableDangerHi = true;
                        DangerHi = (float)(105);
                        EnableAlertHi = true;
                        AlertHi = (float)(60);
                        break;

                    default:
                        AlarmMethod = 0;
                        EnableDangerHi = false;
                        DangerHi = (float)(0);
                        EnableAlertHi = false;
                        AlertHi = (float)(0);
                        break;

                }
                ReturnValue = true;
            }

            GenericTools.DebugMsg("GetScalarStdAlarmLevel(" + PointId.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }




        private static uint GetContainerType_Last_TreeElemId = 0;
        private static uint GetContainerType_Last_Return = 0;
        public uint GetContainerType(uint GetContainerType_TreeElemId)
        {
            if (!IsConnected) return 0;
            if (GetContainerType_TreeElemId == 0) return 0;
            if (GetContainerType_TreeElemId != GetContainerType_Last_TreeElemId)
            {
                GetContainerType_Last_Return = (uint)SQLtoInt("ContainerType", "TreeElem", "TreeElemId=" + GetContainerType_TreeElemId.ToString());
                GetContainerType_Last_TreeElemId = GetContainerType_TreeElemId;
            }
            return GetContainerType_Last_Return;
        }

        /**
        private static uint GetDataId_Last_TreeElemId = 0;
        private static string GetDataId_Last_Result = string.Empty;
        public string GetDataId(uint nGetDataId_TreeElemId)
        {
            if (!IsConnected) return string.Empty;
            if (nGetDataId_TreeElemId < 1) return string.Empty;
            if (nGetDataId_TreeElemId != GetDataId_Last_TreeElemId)
            {
                GetDataId_Last_TreeElemId = nGetDataId_TreeElemId;
                GetDataId_Last_Result = ("1 1 " + GetTblSetId(nGetDataId_TreeElemId).ToString() + " " + GenericTools.AptChangeContainer(GetContainerType(nGetDataId_TreeElemId)).ToString() + " " + nGetDataId_TreeElemId.ToString());
            }
            return GetDataId_Last_Result;
        }

        private static uint GetDataIdPath_Last_TreeElemId = 0;
        private static string GetDataIdPath_Last_Result = string.Empty;
        private static bool GetDataIdPath_Last_Control = true;
        public string GetDataIdPath(uint nGetDataIdPath_TreeElemId, uint nGetDataIdPath_TblSetId)
        {
            if (!IsConnected) return string.Empty;
            if (nGetDataIdPath_TreeElemId < 1)
            {
                return "1 1 " + nGetDataIdPath_TblSetId.ToString() + " 3 0|1 1 " + nGetDataIdPath_TblSetId.ToString() + " 2 0|";
            }
            if (nGetDataIdPath_TreeElemId != GetDataIdPath_Last_TreeElemId)
            {
                if (GetDataIdPath_Last_Control)
                {
                    GetDataIdPath_Last_Control = false;
                    GetDataIdPath_Last_TreeElemId = nGetDataIdPath_TreeElemId;
                    GetDataIdPath_Last_Result = "1 1 " + GetTblSetId(nGetDataIdPath_TreeElemId).ToString() + " " + GenericTools.AptChangeContainer(GetContainerType(nGetDataIdPath_TreeElemId)) + " " + nGetDataIdPath_TreeElemId.ToString() + "|" + GetDataIdPath(GetParentId(nGetDataIdPath_TreeElemId), GetTblSetId(nGetDataIdPath_TreeElemId));
                    GetDataIdPath_Last_Control = true;
                }
                else
                {
                    return "1 1 " + GetTblSetId(nGetDataIdPath_TreeElemId).ToString() + " " + GenericTools.AptChangeContainer(GetContainerType(nGetDataIdPath_TreeElemId)) + " " + nGetDataIdPath_TreeElemId.ToString() + "|" + GetDataIdPath(GetParentId(nGetDataIdPath_TreeElemId), GetTblSetId(nGetDataIdPath_TreeElemId));
                }
            }
            return GetDataIdPath_Last_Result;
        }
        public string GetDataIdPath(uint nGetDataIdPath_TreeElemId)
        {
            return GetDataIdPath(nGetDataIdPath_TreeElemId, GetTblSetId(nGetDataIdPath_TreeElemId));
        }
        **/

        private static uint GetTblSetId_Last_TreeElemId = 0;
        private static uint GetTblSetId_Last_Result = 0;
        public uint GetTblSetId(uint nAnGetTblSetId_TreeElemId)
        {
            if (!IsConnected) return 0;
            if (nAnGetTblSetId_TreeElemId < 1) return 0;
            if (nAnGetTblSetId_TreeElemId != GetTblSetId_Last_TreeElemId)
            {
                GetTblSetId_Last_TreeElemId = nAnGetTblSetId_TreeElemId;
                GetTblSetId_Last_Result = (uint)SQLtoInt("TblSetId", "TreeElem", "TreeElemId=" + nAnGetTblSetId_TreeElemId.ToString());
            }
            return GetTblSetId_Last_Result;
        }

        private static uint GetParentId_Last_TreeElemId = 0;
        private static uint GetParentId_Last_Return = 0;
        public uint GetParentId(uint ParentId_TreeElemId)
        {
            if (!IsConnected) return 0;
            if (ParentId_TreeElemId == 0) return 0;
            if (ParentId_TreeElemId != GetParentId_Last_TreeElemId)
            {
                GetParentId_Last_Return = (uint)SQLtoInt("ParentId", "TreeElem", "TreeElemId=" + ParentId_TreeElemId.ToString());
                GetParentId_Last_TreeElemId = ParentId_TreeElemId;
            }
            return GetParentId_Last_Return;
        }

        private static uint GetParentRefId_Last_TreeElemId = 0;
        private static uint GetParentRefId_Last_Return = 0;
        public uint GetParentRefId(uint GetParentRefId_TreeElemId)
        {
            if (!IsConnected) return 0;
            if (GetParentRefId_TreeElemId == 0) return 0;
            if (GetParentRefId_TreeElemId != GetParentRefId_Last_TreeElemId)
            {
                GetParentRefId_Last_Return = (uint)SQLtoInt("ParentRefId", "TreeElem", "TreeElemId=" + GetParentRefId_TreeElemId.ToString());
                GetParentRefId_Last_TreeElemId = GetParentRefId_TreeElemId;
            }
            return GetParentRefId_Last_Return;
        }

        public DataTable GetTree() { return GetTree(0, 1); }
        public DataTable GetTree(Int32 ParentId) { return GetTree(ParentId, 1); }
        public DataTable GetTree(Int32 ParentId, int HierarchyType)
        {
            if (GetTreeElem_FirstRun) GetTreeElem(false, false);

            string sHierarchyType = string.Empty;
            if (HierarchyType > 0) sHierarchyType = " and HierarchyType=" + HierarchyType.ToString();
            if (HierarchyType > 3) sHierarchyType = " and HierarchyType=1";

            DataView oGetTree_Result_TMP = new DataView(TreeElem, "ParentId=" + ParentId.ToString() + sHierarchyType, (ParentId == 0 ? "Name" : "SlotNumber"), DataViewRowState.CurrentRows);
            DataTable oGetTree_Result = oGetTree_Result_TMP.ToTable();
            //DataRow[] oGetTree_Result = TreeElem.Select();  // RecordSet("select TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + GenericTools.iif(DBVersion.StartsWith("4."), "", ", Overdue, ChannelEnable") + " from " + Owner + "TreeElem where ParentId=" + ParentId.ToString() + sHierarchyType + " order by " + GenericTools.iif(ParentId == 0, "Name", "SlotNumber"));
            int iGetTree_Count = oGetTree_Result.Rows.Count;

            if (iGetTree_Count > 0)
                for (int i = 0; i < iGetTree_Count; i++)
                    if (Convert.ToInt32(oGetTree_Result.Rows[i]["ContainerType"].ToString()) < 4)
                        oGetTree_Result.Merge(GetTree(Convert.ToInt32(oGetTree_Result.Rows[i]["TreeElemId"]), HierarchyType));

            return oGetTree_Result;
        }

        /*
        public AnMeasurement GetMeasurement(ref AnPoint AnPt)
        {
            DateTime StartTime = DateTime.Now;
            int nMeasId = SQLtoInt("max(MeasId)", "MeasReading", "PointId=" + AnPt.TreeElem.TreeElemId.ToString());

            while ((nMeasId <= MeasReadingElements.Length) & (nMeasId > 0))
            {
                if (MeasReadingElements[nMeasId] == null)
                {
                    GenericTools.DebugMsg("GetMeasurement(" + AnPt.TreeElem.TreeElemId + "): " + (DateTime.Now - StartTime).ToString());
                    return GetMeasurement(ref AnPt, nMeasId);
                }
                nMeasId = SQLtoInt("max(MeasId)", "MeasReading", "PointId=" + AnPt.TreeElem.TreeElemId.ToString() + " and MeasId<" + nMeasId.ToString());
            }
            GenericTools.DebugMsg("GetMeasurement(" + AnPt.TreeElem.TreeElemId + "): " + (DateTime.Now - StartTime).ToString());
            if (nMeasId <= 0) return null;
            return GetMeasurement(ref AnPt, nMeasId);
        }
        public AnMeasurement GetMeasurement(ref AnPoint AnPt, Int32 nMeasId)
        {
            DateTime StartTime = DateTime.Now;
            if (nMeasId > MeasReadingElements.Length)
            {
                InitializeMeasReading();
                if (nMeasId > MeasReadingElements.Length)
                {
                    GenericTools.DebugMsg("GetMeasurement(" + AnPt.TreeElem.TreeElemId + ", " + nMeasId + "): " + (DateTime.Now - StartTime).ToString());
                    return null;
                }
            }
            if (MeasReadingElements[nMeasId] == null)
            {
                MeasReadingElements[nMeasId] = new AnMeasurement(ref AnPt);
                MeasReadingElements[nMeasId].ReloadMeas(nMeasId);
            }

            GenericTools.DebugMsg("GetMeasurement(" + AnPt.TreeElem.TreeElemId + ", " + nMeasId + "): " + (DateTime.Now - StartTime).ToString());
            return MeasReadingElements[nMeasId];
        }
        */

        public bool CreateDVVar(uint PointId)
        {
            GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): Starting");
            bool ReturnValue = false;
            if (IsConnected)
            {
                if (PointId > 0)
                {
                    if (GetContainerType(PointId) == 4)
                    {
                        DataTable TreeElem = RecordSet("select TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (DBVersion.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + " from " + Owner + "TreeElem where TreeElemId=" + PointId.ToString());
                        if (TreeElem.Rows.Count > 0)
                        {
                            if (SQLtoInt("count(*)", "TreeElem", "Name='" + TreeElem.Rows[0]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "' and ParentId=" + TreeElem.Rows[0]["ParentId"].ToString()) > 0)
                            {
                                ReturnValue = true;
                                GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): DVVar point exists");
                            }
                            else
                            {
                                if (Convert.ToInt16(TreeElem.Rows[0]["HierarchyType"]) > 1)
                                    PointId = Convert.ToUInt32(TreeElem.Rows[0]["ReferenceId"]);
                                if (PointId > 0)
                                {
                                    int NewTreeElemId = 0;
                                    switch (DBType)
                                    {
                                        case DBType.Oracle: //Oracle
                                            SQLExec("insert into " + Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (DBVersion.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                                + TreeElem.Rows[0]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElem.Rows[0]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElem.Rows[0]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + TreeElem.Rows[0]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElem.Rows[0]["Name"].ToString() + " Variacao', " //Description
                                                + TreeElem.Rows[0]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElem.Rows[0]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + "1, " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElem.Rows[0]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElem.Rows[0]["ParentId"].ToString() + ", " //ParentRefId
                                                + "0, " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (DBVersion.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );
                                            NewTreeElemId = SQLtoInt("select " + Owner + "TreeElemId_Seq.CurrVal from Dual");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Dad_Id") + ", 1, '" + GetRegistrationId("SKFCM_ASDD_DerivedPoint") + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Point_Type_Id") + ", 1, '" + GetRegistrationId("SKFCM_ASPT_DerivedPointCalculated") + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + GetRegistrationId("SKFCM_ASUT_NoUnitsEnglish") + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Location") + ", 3, '')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Application_Id") + ", 1, '" + GetRegistrationId("SKFCM_ASAS_General") + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Owner + "TreeElemId_Seq.CurrVal, " + GetRegistrationId("SKFCM_ASPF_Sensor") + ", 1, '" + SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                            SQLExec("insert into ScalarAlarm (ScalarAlrmId, ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (ScalarAlrmId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80)");
                                            SQLExec("insert into AlarmAssign (AssignId, AlarmId, ElementId, AlarmIndex, Type) values (AssignId_Seq.NextVal, ScalarAlrmId_Seq.CurrVal, TreeElemId_Seq.CurrVal, 1, " + GetRegistrationId("SKFCM_ASAT_Overall") + ")");

                                            SQLExec("insert into DPExpressionDef (ExpressionId, OwnerId, Name) values (DPExpressionDefId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private expression>')");
                                            SQLExec("insert into DPExpressionAssign (AssignId, DPId, ExpId) values (DPExpressionAssignId_Seq.NextVal, TreeElemId_Seq.CurrVal, DPExpressionDefId_Seq.CurrVal)");

                                            SQLExec("insert into DPExpressionVar (VarKey, VarName, VarType, ExpId) values (DPExpressionVarId_Seq.NextVal, 'Var_" + TreeElem.Rows[0]["Name"].ToString().Replace(" ", "_") + "', 1, DPExpressionDefId_Seq.CurrVal)");

                                            SQLExec("insert into DPExpressionVarRef (VarRefId, VarKey, DPId, SourcePtId, ExpId) values (DPExpressionVarRefId_Seq.NextVal, DPExpressionVarId_Seq.CurrVal, TreeElemId_Seq.CurrVal, " + PointId.ToString() + ",  DPExpressionDefId_Seq.CurrVal)");

                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, DPExpressionDefId_Seq.CurrVal, 1002, 3009)");
                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, DPExpressionDefId_Seq.CurrVal, 1001, 2006)");
                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, DPExpressionDefId_Seq.CurrVal, 1005, 1)");
                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, DPExpressionDefId_Seq.CurrVal, 1003, DPExpressionVarId_Seq.CurrVal)");
                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, DPExpressionDefId_Seq.CurrVal, 1001, 2007)");
                                            TreeElem = RecordSet("select TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger " + (DBVersion.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + "from " + Owner + "TreeElem where ReferenceId=" + PointId.ToString());
                                            if (TreeElem.Rows.Count > 0)
                                            {
                                                for (int i = 0; i < TreeElem.Rows.Count; i++)
                                                {
                                                    SQLExec("insert into " + Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (DBVersion.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                                        + TreeElem.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                        + TreeElem.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                        + SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                        + TreeElem.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                        + "'" + TreeElem.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                                        + "4, " //ContainerType
                                                        + "'" + TreeElem.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                                        + TreeElem.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                        + TreeElem.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                        + TreeElem.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                        + "1, " //AlarmFlags
                                                        + TreeElem.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                        + TreeElem.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                        + NewTreeElemId.ToString() + ", " //ReferenceId
                                                        + "0, " //Good
                                                        + "0, " //Alert
                                                        + (DBVersion.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                        );
                                                }
                                            }
                                            break;

                                        case DBType.MSSQL:
                                            NewTreeElemId = SQLtoInt("insert into " + Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (DBVersion.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                                + TreeElem.Rows[0]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElem.Rows[0]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElem.Rows[0]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + TreeElem.Rows[0]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElem.Rows[0]["Name"].ToString() + " Variacao', " //Description
                                                + TreeElem.Rows[0]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElem.Rows[0]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + "1, " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElem.Rows[0]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElem.Rows[0]["ParentId"].ToString() + ", " //ParentRefId
                                                + "0, " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (DBVersion.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                + ";select @@Identity;"
                                                );
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Dad_Id") + ", 1, '" + GetRegistrationId("SKFCM_ASDD_DerivedPoint") + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Point_Type_Id") + ", 1, '" + GetRegistrationId("SKFCM_ASPT_DerivedPointCalculated") + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + GetRegistrationId("SKFCM_ASUT_NoUnitsEnglish") + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Location") + ", 3, '')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + SQLtoString("ValueString", "Point", "FieldId=" + GetRegistrationId("SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Application_Id") + ", 1, '" + GetRegistrationId("SKFCM_ASAS_General") + "')");
                                            SQLExec("insert into " + Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + GetRegistrationId("SKFCM_ASPF_Sensor") + ", 1, '" + SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                            int ScalarAlrmId = SQLtoInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80);select @@Identity;");
                                            SQLExec("insert into AlarmAssign (AlarmId, ElementId, AlarmIndex, Type) values (" + ScalarAlrmId.ToString() + ", " + NewTreeElemId.ToString() + ", 1, " + GetRegistrationId("SKFCM_ASAT_Overall") + ")");

                                            int ExpressionId = SQLtoInt("insert into DPExpressionDef (OwnerId, Name) values (" + NewTreeElemId.ToString() + ", '<Private expression>');select @@Identity;");

                                            SQLExec("insert into DPExpressionAssign (DPId, ExpId) values (" + NewTreeElemId.ToString() + ", " + ExpressionId.ToString() + ")");

                                            int VarKey = SQLtoInt("insert into DPExpressionVar (VarName, VarType, ExpId) values ('Var_" + TreeElem.Rows[0]["Name"].ToString().Replace(" ", "_") + "', 1, " + ExpressionId.ToString() + ");select @@Identity;");

                                            int VarRefId = SQLtoInt("insert into DPExpressionVarRef (VarKey, DPId, SourcePtId, ExpId) values (" + VarKey.ToString() + ", " + NewTreeElemId.ToString() + ", " + PointId.ToString() + ", " + ExpressionId.ToString() + ")");

                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, " + ExpressionId.ToString() + ", 1002, 3009)");
                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, " + ExpressionId.ToString() + ", 1001, 2006)");
                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, " + ExpressionId.ToString() + ", 1005, 1)");
                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, " + ExpressionId.ToString() + ", 1003, " + VarKey.ToString() + ")");
                                            SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, " + ExpressionId.ToString() + ", 1001, 2007)");
                                            TreeElem = RecordSet("select TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger from " + Owner + "TreeElem where ReferenceId=" + PointId.ToString());
                                            if (TreeElem.Rows.Count > 0)
                                            {
                                                for (int i = 0; i < TreeElem.Rows.Count; i++)
                                                {
                                                    SQLExec("insert into " + Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (DBVersion.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                                        + TreeElem.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                        + TreeElem.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                        + SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                        + TreeElem.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                        + "'" + TreeElem.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                                        + "4, " //ContainerType
                                                        + "'" + TreeElem.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                                        + TreeElem.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                        + TreeElem.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                        + TreeElem.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                        + "1, " //AlarmFlags
                                                        + TreeElem.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                        + TreeElem.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                        + NewTreeElemId.ToString() + ", " //ReferenceId
                                                        + "0, " //Good
                                                        + "0, " //Alert
                                                        + (DBVersion.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                        );
                                                }
                                            }
                                            break;
                                    }
                                    ReturnValue = true;
                                }
                            }
                        }
                    }
                }
            }
            GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public bool DeleteTree(uint nTreeElemId)
        {
            bool ReturnValue = false;
            if (IsConnected)
            {
                if (SQLtoInt("count(*)", "TreeElem", "TreeElemId=" + nTreeElemId.ToString()) > 0)
                {
                    int nHierarchyType = SQLtoInt("HierarchyType", "TreeElem", "TreeElemId=" + nTreeElemId.ToString());
                    int nContainerType = SQLtoInt("ContainerType", "TreeElem", "TreeElemId=" + nTreeElemId.ToString());
                    if (nContainerType > 0 & nHierarchyType > 0) ReturnValue = _DeleteTree(nTreeElemId, nHierarchyType, nContainerType);
                }
            }
            GenericTools.DebugMsg("DeleteTree(" + nTreeElemId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        private bool _DeleteTree(uint nTreeElemId, int nHierarchyType, int nContainerType)
        {
            bool ReturnValue = false;
            if (IsConnected)
            {
                ReturnValue = true;
                if (nContainerType < 4)
                {
                    DataTable ChieldList = RecordSet("select TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (DBVersion.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + " from " + Owner + "TreeElem where ParentId=" + nTreeElemId.ToString());
                    if (ChieldList.Rows.Count > 0)
                    {
                        for (int i = 0; i < ChieldList.Rows.Count; i++)
                        {
                            ReturnValue = ReturnValue & _DeleteTree(Convert.ToUInt32(ChieldList.Rows[i]["TreeElemId"]), Convert.ToInt32(ChieldList.Rows[i]["HierarchyType"]), Convert.ToInt32(ChieldList.Rows[i]["ContainerType"]));
                        }
                    }
                }
                else
                {
                    if (nHierarchyType == 1)
                    {
                        ReturnValue = ReturnValue & SQLExec("delete from " + Owner + "MeasReading where PointId=" + nTreeElemId.ToString());
                        ReturnValue = ReturnValue & SQLExec("delete from " + Owner + "Measurement where PointId=" + nTreeElemId.ToString());
                        ReturnValue = ReturnValue & SQLExec("delete from " + Owner + "TreeElem where ReferenceId=" + nTreeElemId.ToString());
                    }
                }
                if (nContainerType > 1) ReturnValue = ReturnValue & SQLExec("delete from " + Owner + "TreeElem where TreeElemId=" + nTreeElemId.ToString());
            }
            GenericTools.DebugMsg("DeleteTree(" + nTreeElemId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public string GetAlarmLevelText(int nAlarmLevel)
        {
            string ReturnValue = nAlarmLevel.ToString();

            switch (nAlarmLevel)
            {
                case 0:
                    ReturnValue = "Invalid";
                    break;

                case 1:
                    ReturnValue = "None";
                    break;

                case 2:
                    ReturnValue = "Normal";
                    break;

                case 3:
                    ReturnValue = "Alert";
                    break;

                case 4:
                    ReturnValue = "Danger";
                    break;
            }
            return ReturnValue;
        }

        private static Int32 GetParentTree_Last_TreeElemId = 0;
        private static string GetParentTree_Last_Result = string.Empty;
        public string GetParentTree(Int32 nTreeElemId)
        {
            if (nTreeElemId == GetParentTree_Last_TreeElemId) return GetParentTree_Last_Result;
            GetParentTree_Last_TreeElemId = nTreeElemId;
            GetParentTree_Last_Result = _GetParentTree(nTreeElemId);
            return GetParentTree_Last_Result;
        }
        private string _GetParentTree(Int32 nTreeElemId)
        {
            //if (nTreeElemId==GetParentTree_Last_TreeElemId) return GetParentTree_Last_Result;

            string ResultValue = string.Empty;
            int nParentId = SQLtoInt("ParentId", "TreeElem", "TreeElemId=" + nTreeElemId.ToString());
            string sName = SQLtoString("Name", "TreeElem", "TreeElemId=" + nTreeElemId.ToString());
            if (nParentId > 0)
            {
                return GetParentTree(nParentId) + "\\" + sName;
            }
            else if (nParentId == 0)
            {
                sName = SQLtoString("TblSetName", "TableSet", "TblSetId=" + SQLtoString("TblSetId", "TreeElem", "TreeElemId=" + nTreeElemId.ToString()));
            }
            return sName;
        }
    }

    public class TableSet
    {
        private bool _IsLoaded = false;
        public bool IsLoaded { get { return _IsLoaded; } }

        private AnalystConnection _Connection = null;
        public AnalystConnection Connection { get { return _Connection; } set { _Connection = value; } }

        private uint _TblSetId;
        public uint TblSetId { get { return (IsLoaded ? _TblSetId : 0); } }

        private string _TblSetName;
        public string TblSetName { get { return (IsLoaded ? _TblSetName : null); } }

        private uint _CustomerId;
        public uint CustomerId { get { return (IsLoaded ? _CustomerId : 0); } }

        public TableSet() { }
        public TableSet(AnalystConnection AnalystConnection)
        {
            Connection = AnalystConnection;
        }
        public TableSet(AnalystConnection AnalystConnection, uint TableSetId)
        {
            Load(AnalystConnection, TableSetId);
        }
        public TableSet(AnalystConnection AnalystConnection, string TableSetName)
        {
            Load(AnalystConnection, TableSetName);
        }

        public List<TableSet> SelectAll(string Where = null)
        {
            List<TableSet> _Select = new List<TableSet>();
            DataTable _dSelect = Connection.DataTable("SELECT * FROM TABLESET");

            foreach (DataRow item in _dSelect.Rows)
            {
                _Select.Add(new TableSet(Connection, uint.Parse(item["TblSetId"].ToString())));

            }

            return _Select;
        }

        public TableSet Select(uint TblSetId)
        {
            TableSet _Select = new TableSet(Connection, TblSetId);
            return _Select;
        }

        public bool Load(AnalystConnection AnalystConnection, uint TableSetId)
        {
            _IsLoaded = false;

            if (AnalystConnection.IsConnected)
            {
                DataTable TableSet = AnalystConnection.DataTable("TblSetId, TblSetName, CustomerId", "TableSet", "TblSetId=" + TableSetId.ToString());
                if (TableSet.Rows.Count > 0)
                {
                    _TblSetId = Convert.ToUInt32(TableSet.Rows[0]["TblSetId"]);
                    _TblSetName = TableSet.Rows[0]["TblSetName"].ToString();
                    _CustomerId = Convert.ToUInt32(TableSet.Rows[0]["CustomerId"]);

                    _IsLoaded = true;
                }
            }

            return _IsLoaded;
        }
        public bool Load(AnalystConnection AnalystConnection, string TableSetName)
        {
            _IsLoaded = false;

            if (AnalystConnection.IsConnected)
            {
                DataTable TableSet = AnalystConnection.DataTable("TblSetId, TblSetName, CustomerId", "TableSet", "upper(TblSetName)='" + TableSetName.ToUpper() + "'");
                if (TableSet.Rows.Count > 0)
                {
                    _TblSetId = (uint)TableSet.Rows[0]["TblSetId"];
                    _TblSetName = TableSet.Rows[0]["TblSetName"].ToString();
                    _CustomerId = (uint)TableSet.Rows[0]["CustomerId"];

                    _IsLoaded = true;
                }
            }

            return _IsLoaded;
        }
    }

    /// <summary>
    /// Deprecated
    /// </summary>
    //public class AnTreeElem : TreeElem
    //{
    //}
    public class TreeElem
    {
        //private bool _IsLoaded = false;
        public bool IsLoaded { get { return (Connection.IsConnected && (TreeElemId > 0)); } }

        private AnalystConnection _Connection;
        public AnalystConnection Connection
        {
            get { return _Connection; }
            set
            {
                _TreeElemRow = null;
                _Connection = value;
            }
        }


        private DataRow _TreeElemRow = null;
        public DataRow TreeElemRow
        {
            get
            {
                if (_TreeElemRow == null)
                {
                    if (IsLoaded)
                    {
                        DataTable TreeElemRows = Connection.DataTable("*", "TreeElem", "TreeElemId=" + TreeElemId.ToString());
                        if (TreeElemRows.Rows.Count > 0)
                            _TreeElemRow = TreeElemRows.Rows[0];
                        else
                            _TreeElemRow = null;
                    }
                    else
                        _TreeElemRow = null;
                }
                return _TreeElemRow;
            }
        }

        private DataTable _Tree = null;
        public DataTable Tree
        {
            get
            {
                if (_Tree == null)
                {
                    _Tree = new DataTable();
                    if (ContainerType != ContainerType.Point)
                        foreach (TreeElem ChildItem in Child)
                            _Tree.Merge(ChildItem.Tree);
                    _Tree.ImportRow(TreeElemRow);
                }
                return _Tree;
            }
        }

        private uint _TreeElemId;
        public uint TreeElemId
        {
            get
            {
                return _TreeElemId;
            }
            set
            {
                if (_TreeElemId != value)
                {

                    _DataId = null;
                    _DataIdPath = null;
                    _Child = null;
                    _Parent = null;
                    _ParentRef = null;
                    _TableSet = null;
                    _Tree = null;
                    _TreeElemRow = null;

                    _TreeElemId = value;
                }
            }
        }

        public uint HierarchyId
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["HierarchyId"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != HierarchyId)
                    if (Connection.SQLUpdate("TreeElem", "HierarchyId", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public uint BranchLevel
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["BranchLevel"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != BranchLevel)
                    if (Connection.SQLUpdate("TreeElem", "BranchLevel", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public uint SlotNumber
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["SlotNumber"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != SlotNumber)
                    if (Connection.SQLUpdate("TreeElem", "SlotNumber", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public uint TblSetId
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["TblSetId"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != TblSetId)
                    if (Connection.SQLUpdate("TreeElem", "TblSetId", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public string Name
        {
            get
            {
                return (IsLoaded ? TreeElemRow["Name"].ToString() : string.Empty);
            }
            set
            {
                if (IsLoaded)
                    //if (value != Name)
                    if (Connection.SQLUpdate("TreeElem", "Name", value, "TreeElemId=" + TreeElemId.ToString() + " or ReferenceId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public ContainerType ContainerType
        {
            get
            {
                return (IsLoaded ? (Analyst.ContainerType)Convert.ToUInt32(TreeElemRow["ContainerType"]) : Analyst.ContainerType.None);
            }
            set
            {
                if (IsLoaded)
                    //if (value != ContainerType)
                    if (Connection.SQLUpdate("TreeElem", "ContainerType", (uint)value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public string Description
        {
            get
            {
                return (IsLoaded ? TreeElemRow["Description"].ToString() : string.Empty);
            }
            set
            {
                if (IsLoaded)
                    //if (value != Description)
                    if (_Connection.SQLUpdate("TreeElem", "Description", value, "TreeElemId=" + TreeElemId.ToString() + " or ReferenceId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public bool ElementEnable
        {
            get
            {
                return (IsLoaded ? (Convert.ToUInt32(TreeElemRow["ElementEnable"]) == 1) : false);
            }
            set
            {
                if (IsLoaded)
                    //if (value != ElementEnable)
                    if (Connection.SQLUpdate("TreeElem", "ElementEnable", Convert.ToUInt32(value), "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public bool ParentEnable
        {
            get
            {
                return (IsLoaded ? (Convert.ToUInt32(TreeElemRow["ParentEnable"]) == 1) : false);
            }
            set
            {
                if (IsLoaded)
                    //if (value != ParentEnable)
                    if (Connection.SQLUpdate("TreeElem", "ParentEnable", Convert.ToUInt32(value), "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public HierarchyType HierarchyType
        {
            get
            {
                return (IsLoaded ? (Analyst.HierarchyType)Convert.ToUInt32(TreeElemRow["HierarchyType"]) : Analyst.HierarchyType.None);
            }
            set
            {
                if (IsLoaded)
                    //if (value != HierarchyType)
                    if (Connection.SQLUpdate("TreeElem", "HierarchyType", (uint)value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public AlarmFlags AlarmFlags
        {
            get
            {
                return (IsLoaded ? (Analyst.AlarmFlags)Convert.ToUInt32(TreeElemRow["AlarmFlags"]) : Analyst.AlarmFlags.None);
            }
            set
            {
                if (IsLoaded)
                    //if (value != AlarmFlags)
                    if (Connection.SQLUpdate("TreeElem", "AlarmFlags", (uint)value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }

        public uint ParentId
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["ParentId"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != ParentId)
                    if (Connection.SQLUpdate("TreeElem", "ParentId", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public uint ParentRefId
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["ParentRefId"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != ParentRefId)
                    if (Connection.SQLUpdate("TreeElem", "ParentRefId", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public uint ReferenceId
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["ReferenceId"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != ReferenceId)
                    if (Connection.SQLUpdate("TreeElem", "ReferenceId", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public uint Good
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["Good"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value!=Good)
                    if (Connection.SQLUpdate("TreeElem", "Good", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public uint Alert
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["Alert"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != Alert)
                    if (Connection.SQLUpdate("TreeElem", "Alert", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public uint Danger
        {
            get
            {
                return (IsLoaded ? Convert.ToUInt32(TreeElemRow["Danger"]) : 0);
            }
            set
            {
                if (IsLoaded)
                    //if (value != Danger)
                    if (Connection.SQLUpdate("TreeElem", "Danger", value, "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public bool Overdue
        {
            get
            {
                return (IsLoaded ? (Convert.ToUInt32(TreeElemRow["Overdue"]) == 1) : false);
            }
            set
            {
                if (IsLoaded)
                    //if (value != Overdue)
                    if (Connection.SQLUpdate("TreeElem", "Overdue", Convert.ToUInt32(value), "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }
        public bool ChannelEnable
        {
            get
            {
                return (IsLoaded ? (Convert.ToUInt32(TreeElemRow["ChannelEnable"]) == 1) : false);
            }
            set
            {
                if (IsLoaded)
                    //if (value != ChannelEnable)
                    if (Connection.SQLUpdate("TreeElem", "ChannelEnable", Convert.ToUInt32(value), "TreeElemId=" + TreeElemId.ToString()) > 0)
                        _TreeElemRow = null;
            }
        }

        private string _DataId = null;
        public string DataId
        {
            get
            {
                if (string.IsNullOrEmpty(_DataId) && Connection.IsConnected)
                    _DataId = ("1 1 " + TblSetId.ToString() + " " + GenericTools.AptChangeContainer((uint)ContainerType).ToString() + " " + TreeElemId.ToString());
                return _DataId;
            }
        }

        private string _DataIdPath = null;
        public string DataIdPath
        {
            get
            {
                if (string.IsNullOrEmpty(_DataIdPath) && Connection.IsConnected)
                    _DataIdPath = "1 1 " + TblSetId.ToString() + " " + GenericTools.AptChangeContainer((uint)ContainerType) + " " + TreeElemId.ToString() + ((ContainerType == Analyst.ContainerType.Root) ? "" : "|" + Parent.DataIdPath);
                return _DataIdPath;
            }
        }

        private TableSet _TableSet = null;
        public TableSet TableSet
        {
            get
            {
                _TableSet = (IsLoaded ? ((_TableSet == null) ? new TableSet(Connection, TblSetId) : (_TableSet.TblSetId != TblSetId ? new TableSet(Connection, TblSetId) : _TableSet)) : null);
                return _TableSet;
            }
        }

        private TreeElem _Parent = null;
        public TreeElem Parent
        {
            get
            {
                if ((ParentId < 2147000000) & (ContainerType != ContainerType.Root) & (_Parent == null))
                    _Parent = new TreeElem(Connection, ParentId);
                return _Parent;
            }
        }

        private TreeElem _ParentRef = null;
        public TreeElem ParentRef
        {
            get
            {
                if ((ParentRefId < 2147000000) & (ContainerType != ContainerType.Root) & (_ParentRef == null))
                    _ParentRef = new TreeElem(Connection, ParentRefId);
                return _ParentRef;
            }
        }

        private List<TreeElem> _Child = null; //new List<TreeElem>();
        public List<TreeElem> Child
        {
            get
            {
                try
                {
                    if (IsLoaded)
                    {
                        if ((ContainerType == ContainerType.Set) | (ContainerType == ContainerType.Machine))
                        {
                            if (_Child == null)
                            {
                                DataTable TempDataTable = Connection.DataTable("TreeElemId", "TreeElem", "ParentId=" + TreeElemId.ToString() + " order by SlotNumber asc");
                                if (TempDataTable.Rows.Count > 0)
                                {
                                    _Child = new List<TreeElem>();
                                    for (int i = 0; i < TempDataTable.Rows.Count; i++)
                                        _Child.Add(new TreeElem(Connection, Convert.ToUInt32(TempDataTable.Rows[i]["TreeElemId"])));
                                }
                                else
                                    _Child = null; //new List<TreeElem>();
                            }
                        }
                        else
                            _Child = null; //new List<TreeElem>();
                    }
                }
                catch (Exception ex)
                {
                    GenericTools.DebugMsg("TreeElem.Child error: " + ex.Message);
                }
                return _Child;
            }
        }
        

        public TreeElem(AnalystConnection AnalystConnection, uint TreeElemId)
        {
            GenericTools.DebugMsg("TreeElem(" + TreeElemId.ToString() + "): Starting...");

            try
            {
                Connection = AnalystConnection;
                this.TreeElemId = TreeElemId;
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("TreeElem(" + TreeElemId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("TreeElem(" + TreeElemId.ToString() + "): Finished!");
        }
        public TreeElem(AnalystConnection AnalystConnection,
            uint HierarchyId,
            uint BranchLevel,
            uint SlotNumber,
            uint TblSetId,
            string Name,
            ContainerType ContainerType,
            string Description,
            bool ElementEnable,
            bool ParentEnable,
            uint HierarchyType,
            AlarmFlags AlarmFlags,
            uint ParentId,
            uint ParentRefId,
            uint ReferenceId,
            uint Good,
            uint Alert,
            uint Danger,
            bool Overdue,
            bool ChannelEnable)
        {
            GenericTools.DebugMsg("TreeElem(" + Name + "): Starting...");

            try
            {
                _Connection = AnalystConnection;
                TreeElemId = Add(HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger, Overdue, ChannelEnable).TreeElemId;
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("TreeElem(" + Name + ") error: " + ex.Message);
                TreeElemId = 0;
            }

            GenericTools.DebugMsg("TreeElem(" + Name + "): Finished!");
        }

        public TreeElem Add(
            uint HierarchyId,
            uint BranchLevel,
            uint SlotNumber,
            uint TblSetId,
            string Name,
            ContainerType ContainerType,
            string Description,
            bool ElementEnable,
            bool ParentEnable,
            uint HierarchyType,
            AlarmFlags AlarmFlags,
            uint ParentId,
            uint ParentRefId,
            uint ReferenceId,
            uint Good,
            uint Alert,
            uint Danger,
            bool Overdue,
            bool ChannelEnable)
        {
            GenericTools.DebugMsg("TreeElem.Add(" + Name + "): Starting...");

            TreeElem ReturnValue = null;
            try
            {
                List<TableColumn> Columns = new List<TableColumn>();

                Columns.Add(new TableColumn("TreeElemId", 0));
                Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                Columns.Add(new TableColumn("BranchLevel", BranchLevel));
                Columns.Add(new TableColumn("SlotNumber", SlotNumber));
                Columns.Add(new TableColumn("TblSetId", TblSetId));
                Columns.Add(new TableColumn("Name", Name));
                Columns.Add(new TableColumn("ContainerType", (uint)ContainerType));
                Columns.Add(new TableColumn("Description", Description));
                Columns.Add(new TableColumn("ElementEnable", Convert.ToUInt32(ElementEnable)));
                Columns.Add(new TableColumn("ParentEnable", Convert.ToUInt32(ParentEnable)));
                Columns.Add(new TableColumn("HierarchyType", (uint)HierarchyType));
                Columns.Add(new TableColumn("AlarmFlags", (uint)AlarmFlags));
                Columns.Add(new TableColumn("ParentId", ParentId));
                Columns.Add(new TableColumn("ParentRefId", ParentRefId));
                Columns.Add(new TableColumn("ReferenceId", ReferenceId));
                Columns.Add(new TableColumn("Good", Good));
                Columns.Add(new TableColumn("Alert", Alert));
                Columns.Add(new TableColumn("Danger", Danger));
                Columns.Add(new TableColumn("Overdue", Convert.ToUInt32(Overdue)));
                Columns.Add(new TableColumn("ChannelEnable", Convert.ToUInt32(ChannelEnable)));

                uint TreeElemIdTMP = Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_seq"));

                if (TreeElemIdTMP > 0)
                {
                    if (ContainerType == Analyst.ContainerType.Set || ContainerType == Analyst.ContainerType.Machine)
                    {
                        Connection.SQLExec("insert into " + Connection.Owner + "GroupTbl (ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority) VALUES " +
                                            "(" + TreeElemIdTMP + ", 0, '','','','','', '','',0)");
                    }
                    ReturnValue = new TreeElem(Connection, TreeElemIdTMP);
                }
            }
            catch (Exception ex)
            {
                ReturnValue = null;
                GenericTools.DebugMsg("TreeElem(" + Name + ") error: " + ex.Message);
            }

            if (ReturnValue == null)
                GenericTools.DebugMsg("TreeElem(" + Name + "): null");
            else
                GenericTools.DebugMsg("TreeElem(" + Name + "): " + ReturnValue.TreeElemId.ToString());

            return ReturnValue;
        }
        #region Asset Name - GROUPTBL
        private string _AssetName = null;
        public string AssetName
        {
            get
            {
                if (_AssetName == null)
                    _AssetName = Connection.SQLtoString("AssetName", "GroupTbl", "ElementId=" + TreeElemId.ToString());
                return _AssetName;
            }
            set
            {
                if (Connection.SQLUpdate("GroupTbl", "AssetName", value, "ElementId=" + TreeElemId.ToString()) > 0)
                    _AssetName = value;
                else
                    _AssetName = null;
            }
        }
        private string _SegmentName = null;
        public string SegmentName
        {
            get
            {
                if (_SegmentName == null)
                    _SegmentName = Connection.SQLtoString("SegmentName", "GroupTbl", "ElementId=" + TreeElemId.ToString());
                return _SegmentName;
            }
            set
            {
                if (Connection.SQLUpdate("GroupTbl", "SegmentName", value, "ElementId=" + TreeElemId.ToString()) > 0)
                    _SegmentName = value;
                else
                    _SegmentName = null;
            }
        }
        #endregion

        public Note AddNotes(string Text) { return AddNotes(Text, GenericTools.DateTime()); }
        /// <summary>Add notes to element</summary>
        /// <param name="OwnerId">Element unique id (TreeElemId)</param>
        /// <param name="Text">Notes text</param>
        /// <param name="TimeStamp">Time stamp</param>
        /// <returns>Notes unique id</returns>
        public Note AddNotes(string Text, DateTime TimeStamp) { return AddNotes(Text, GenericTools.DateTime(TimeStamp)); }
        /// <summary>Add notes to element</summary>
        /// <param name="OwnerId">Element unique id (TreeElemId)</param>
        /// <param name="Text">Notes text</param>
        /// <param name="DataDtg">Time stamp (YYYYMMDDHHMISS)</param>
        /// <returns>Notes unique id</returns>
        public Note AddNotes(string Text, string DataDtg)
        {
            uint OwnerId = TreeElemId;

            GenericTools.DebugMsg("AddNotes(" + OwnerId.ToString() + ", '" + Text + "', '" + DataDtg + "'): Starting");
            uint ReturnValue = 0;

            try
            {
                if (Connection.IsConnected)
                {
                    ReturnValue = Connection.SQLtoUInt("count(*)", "Notes", "OwnerId=" + OwnerId.ToString() + " and DataDtg='" + DataDtg + "' and Text='" + Text + "'");
                    if (ReturnValue < 1)
                    {
                        switch (Connection.DBType)
                        {
                            case DBType.Oracle: // Oracle
                                Connection.SQLExec("insert into " + Connection.Owner + "Notes (NotesId, OwnerId, Text, DataDtg, Category, AppData, CategoryId, SubCategoryId) values (" +
                                    "NotesId_Seq.NextVal," + //NOTESID
                                    OwnerId.ToString() + "," + //OWNERID
                                    "'" + Text + "'," + //TEXT
                                    "'" + DataDtg + "'," + //DATADTG
                                    "null," + //CATEGORY
                                    "null," + //APPDATA
                                    "0," + //CATEGORYID
                                    "0" + //SUBCATEGORYID
                                    ")");
                                ReturnValue = Connection.SQLtoUInt("select " + Connection.Owner + "NotesId_Seq.CurrVal from Dual");
                                break;

                            case DBType.MSSQL: // MSSQL
                                ReturnValue = Connection.SQLtoUInt("insert into " + Connection.Owner + "Notes (OwnerId, Text, DataDtg, Category, AppData, CategoryId, SubCategoryId) values (" +
                                    OwnerId.ToString() + "," + //OWNERID
                                    "'" + Text + "'," + //TEXT
                                    "'" + DataDtg + "'," + //DATADTG
                                    "null," + //CATEGORY
                                    "null," + //APPDATA
                                    "0," + //CATEGORYID
                                    "0" + //SUBCATEGORYID
                                    ");select @@Identity;");
                                break;

                            default:
                                break;
                        }
                    }
                    else
                    {
                        ReturnValue = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("AddNotes(" + OwnerId.ToString() + ", '" + Text + "', '" + DataDtg + "') error: " + ex.Message);
            }

            //GenericTools.DebugMsg("AddNotes(" + OwnerId.ToString() + ", '" + Text + "', '" + DataDtg + "'): " + ReturnValue.ToString());

            Note Nota = new Note(Connection, ReturnValue);
            return Nota;
        }
        public bool Delete(AnalystConnection Connection, uint TreeElemId, bool CleanEmptyParent = false, bool Delete_Permanent = false)
        {
            GenericTools.DebugMsg("TreeElem.Delete(" + TreeElemId.ToString() + ", " + CleanEmptyParent.ToString() + "): Starting...");


            TreeElem objectTreeElem = new TreeElem(Connection, TreeElemId);
            TreeElem objectParent = null;

            if (objectTreeElem.Parent != null)
                objectParent = objectTreeElem.Parent;

            bool ReturnValue = true;

            try
            {
                //Delete Childs
                if (ContainerType != Analyst.ContainerType.Point)
                {
                    if (objectTreeElem.Child != null)
                    {
                        foreach (TreeElem ChildItem in objectTreeElem.Child)
                        {
                            ReturnValue = (ReturnValue & Delete(Connection, ChildItem.TreeElemId));
                            Console.Write(ReturnValue);
                        }
                    }
                }
                //Delete element
                if (Delete_Permanent == true)
                    ReturnValue = (ReturnValue & (Connection.SQLExec("DELETE FROM TreeElem WHERE TreeElemId=" + TreeElemId.ToString())));
                else
                    ReturnValue = (ReturnValue & (Connection.SQLUpdate("TreeElem", "ParentId", 2147000000, "TreeElemId=" + TreeElemId.ToString()) > 0));


                //Update alarm flags
                if (ReturnValue)
                    objectParent.CalcAlarm();
                else
                    this.CalcAlarm();

                //Delete parents
                if (ContainerType != ContainerType.Root)
                    if ((CleanEmptyParent) && ((objectParent.Child == null) || (Parent.Child.Count < 1)))
                        objectParent.Delete(CleanEmptyParent);

            }
            catch (Exception ex)
            {
                ReturnValue = false;
                GenericTools.DebugMsg("TreeElem.Delete(" + TreeElemId.ToString() + ", " + CleanEmptyParent.ToString() + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("TreeElem.Delete(" + TreeElemId.ToString() + ", " + CleanEmptyParent.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }


        public bool Delete(bool CleanEmptyParent = false, bool Delete_Permanent = false)
        {
            GenericTools.DebugMsg("TreeElem.Delete(" + TreeElemId.ToString() + ", " + CleanEmptyParent.ToString() + "): Starting...");

            TreeElem TmpParent = null;
            if (this.Parent != null)
                TmpParent = Parent;

            bool ReturnValue = true;

            try
            {
                //Delete Childs
                if (ContainerType != Analyst.ContainerType.Point)
                    foreach (TreeElem ChildItem in Child)
                    {
                        ReturnValue = (ReturnValue & ChildItem.Delete());
                        Console.Write(ReturnValue);
                    }
                //Delete element
                if (Delete_Permanent == true)
                    ReturnValue = (ReturnValue & (Connection.SQLExec("DELETE FROM TreeElem WHERE TreeElemId=" + TreeElemId.ToString())));
                else
                    ReturnValue = (ReturnValue & (Connection.SQLUpdate("TreeElem", "ParentId", 2147000000, "TreeElemId=" + TreeElemId.ToString()) > 0));


                //Update alarm flags
                if (ReturnValue && TmpParent != null)
                    TmpParent.CalcAlarm();
                else
                    this.CalcAlarm();

                //Delete parents
                if (ContainerType != ContainerType.Root)
                    if ((CleanEmptyParent) && ((TmpParent.Child == null) || (Parent.Child.Count < 1)))
                        TmpParent.Delete(CleanEmptyParent);

            }
            catch (Exception ex)
            {
                ReturnValue = false;
                GenericTools.DebugMsg("TreeElem.Delete(" + TreeElemId.ToString() + ", " + CleanEmptyParent.ToString() + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("TreeElem.Delete(" + TreeElemId.ToString() + ", " + CleanEmptyParent.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public void CalcAlarm(bool IncludeChild = false, bool FirstCall = true)
        {
            GenericTools.DebugMsg("TreeElem.CalcAlarm(" + TreeElemId.ToString() + ", " + IncludeChild.ToString() + ", " + FirstCall.ToString() + "): Starting...");
            try
            {
                uint tGood = Good;
                uint tAlert = Alert;
                uint tDanger = Danger;
                uint tAlarmFlags = (uint)AlarmFlags;

                if (ContainerType == Analyst.ContainerType.Point)
                {
                    tGood = Connection.SQLtoUInt("count(*)", "MeasAlarm", "PointId=" + TreeElemId.ToString() + " and AlarmLevel=" + ((uint)AlarmFlags.Good).ToString());
                    tAlert = Connection.SQLtoUInt("count(*)", "MeasAlarm", "PointId=" + TreeElemId.ToString() + " and AlarmLevel=" + ((uint)AlarmFlags.Alert).ToString());
                    tDanger = Connection.SQLtoUInt("count(*)", "MeasAlarm", "PointId=" + TreeElemId.ToString() + " and AlarmLevel=" + ((uint)AlarmFlags.Danger).ToString());
                    tAlarmFlags = Connection.SQLtoUInt("max(AlarmLevel)", "MeasAlarm", "PointId=" + TreeElemId.ToString());
                }
                else
                {
                    if (IncludeChild)
                        foreach (TreeElem ChildElement in Child)
                            ChildElement.CalcAlarm(IncludeChild, false);

                    tGood = Connection.SQLtoUInt("sum(Good)", "TreeElem", "ParentId=" + TreeElemId.ToString() + " and ElementEnable=1 and ParentEnable=0");
                    tAlert = Connection.SQLtoUInt("sum(Alert)", "TreeElem", "ParentId=" + TreeElemId.ToString() + " and ElementEnable=1 and ParentEnable=0");
                    tDanger = Connection.SQLtoUInt("sum(Danger)", "TreeElem", "ParentId=" + TreeElemId.ToString() + " and ElementEnable=1 and ParentEnable=0");
                    tAlarmFlags = Connection.SQLtoUInt("max(AlarmFlags)", "TreeElem", "ParentId=" + TreeElemId.ToString() + " and ElementEnable=1 and ParentEnable=0");
                }
                if (((tGood != Good) | (tAlert != Alert) | (tDanger != Danger) | (tAlarmFlags != (uint)AlarmFlags)))
                {
                    Good = tGood;
                    Alert = tAlert;
                    Danger = tDanger;
                    AlarmFlags = (Analyst.AlarmFlags)tAlarmFlags;
                    if (FirstCall) Parent.CalcAlarm();
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("TreeElem.CalcAlarm(" + TreeElemId.ToString() + ", " + IncludeChild.ToString() + ", " + FirstCall.ToString() + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("TreeElem.CalcAlarm(" + TreeElemId.ToString() + ", " + IncludeChild.ToString() + ", " + FirstCall.ToString() + "): Finished!");

        }

        public TreeElem FindTreeElem(string SetName)
        {
            uint teId = Connection.SQLtoUInt("SELECT TREEELEMID FROM TREEELEM WHERE NAME='" + SetName + "' and TblSetId=" + TblSetId);
            TreeElem TreeElem = new TreeElem(Connection, teId);
            return TreeElem;
        }

    }

    public class Set
    {
        private TreeElem _TreeElem;
        public TreeElem TreeElem { get { return _TreeElem; } }

        public Set(TreeElem TreeElem)
        {
            GenericTools.DebugMsg("Set(" + TreeElem.TreeElemId.ToString() + "): Starting...");
            try
            {
                if ((TreeElem.ContainerType == ContainerType.Set) | (TreeElem.ContainerType == ContainerType.Root))
                    _TreeElem = TreeElem;
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Set(" + TreeElem.TreeElemId.ToString() + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("Set(" + TreeElem.TreeElemId.ToString() + "): Finished!");
        }
        public Set(AnalystConnection AnalystConnection, uint TreeElemId)
        {
            GenericTools.DebugMsg("Set(" + TreeElemId.ToString() + "): Starting...");
            TreeElem TreeElemTemp = new TreeElem(AnalystConnection, TreeElemId);
            try
            {
                if ((TreeElemTemp.ContainerType == ContainerType.Set) | (TreeElemTemp.ContainerType == ContainerType.Root))
                    _TreeElem = TreeElemTemp;
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Set(" + TreeElemId.ToString() + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("Set(" + TreeElemId.ToString() + "): Finished!");
        }
    }

    public class Machine
    {

        private AnalystConnection _AnalystConnection = null;
        public AnalystConnection Connection
        {
            get
            {
                if (_AnalystConnection == null)
                    _AnalystConnection = TreeElem.Connection;
                return _AnalystConnection;
            }
            set
            {
                //  if (SetValueString("SKFCM_ASPF_Detection", Convert.ToUInt32(value).ToString()))
                _AnalystConnection = value;
            }
        }


        private TreeElem _TreeElem;
        public TreeElem TreeElem { get { return _TreeElem; } }
        public uint TreeElemId { get { return TreeElem.TreeElemId; } }
        // public AnalystConnection Connection { get { return TreeElem.Connection; } set { Connection = value; } }


        public DataTable SelectAll(string Where = null)
        {
            DataTable Selected = Connection.DataTable("TREEELEMID, NAME, TBLSETID", "TREEELEM", "ContainerType=3 and hierarchytype=1 and Parentid!=2147000000 " + (Where != null ? " AND " + Where : ""));
            return Selected;
        }

        public Note AddNotes(string Text) { return AddNotes(Text, GenericTools.DateTime()); }
        /// <summary>Add notes to element</summary>
        /// <param name="OwnerId">Element unique id (TreeElemId)</param>
        /// <param name="Text">Notes text</param>
        /// <param name="TimeStamp">Time stamp</param>
        /// <returns>Notes unique id</returns>
        public Note AddNotes(string Text, DateTime TimeStamp) { return AddNotes(Text, GenericTools.DateTime(TimeStamp)); }
        /// <summary>Add notes to element</summary>
        /// <param name="OwnerId">Element unique id (TreeElemId)</param>
        /// <param name="Text">Notes text</param>
        /// <param name="DataDtg">Time stamp (YYYYMMDDHHMISS)</param>
        /// <returns>Notes unique id</returns>
        public Note AddNotes(string Text, string DataDtg)
        {
            uint OwnerId = TreeElemId;

            GenericTools.DebugMsg("AddNotes(" + OwnerId.ToString() + ", '" + Text + "', '" + DataDtg + "'): Starting");
            uint ReturnValue = 0;

            try
            {
                if (Connection.IsConnected)
                {
                    ReturnValue = Connection.SQLtoUInt("count(*)", "Notes", "OwnerId=" + OwnerId.ToString() + " and DataDtg='" + DataDtg + "' and Text='" + Text + "'");
                    if (ReturnValue < 1)
                    {
                        switch (Connection.DBType)
                        {
                            case DBType.Oracle: // Oracle
                                Connection.SQLExec("insert into " + Connection.Owner + "Notes (NotesId, OwnerId, Text, DataDtg, Category, AppData, CategoryId, SubCategoryId) values (" +
                                    "NotesId_Seq.NextVal," + //NOTESID
                                    OwnerId.ToString() + "," + //OWNERID
                                    "'" + Text + "'," + //TEXT
                                    "'" + DataDtg + "'," + //DATADTG
                                    "null," + //CATEGORY
                                    "null," + //APPDATA
                                    "0," + //CATEGORYID
                                    "0" + //SUBCATEGORYID
                                    ")");
                                ReturnValue = Connection.SQLtoUInt("select " + Connection.Owner + "NotesId_Seq.CurrVal from Dual");
                                break;

                            case DBType.MSSQL: // MSSQL
                                ReturnValue = Connection.SQLtoUInt("insert into " + Connection.Owner + "Notes (OwnerId, Text, DataDtg, Category, AppData, CategoryId, SubCategoryId) values (" +
                                    OwnerId.ToString() + "," + //OWNERID
                                    "'" + Text + "'," + //TEXT
                                    "'" + DataDtg + "'," + //DATADTG
                                    "null," + //CATEGORY
                                    "null," + //APPDATA
                                    "0," + //CATEGORYID
                                    "0" + //SUBCATEGORYID
                                    ");select @@Identity;");
                                break;

                            default:
                                break;
                        }
                    }
                    else
                    {
                        ReturnValue = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("AddNotes(" + OwnerId.ToString() + ", '" + Text + "', '" + DataDtg + "') error: " + ex.Message);
            }

            //GenericTools.DebugMsg("AddNotes(" + OwnerId.ToString() + ", '" + Text + "', '" + DataDtg + "'): " + ReturnValue.ToString());

            Note Nota = new Note(Connection, ReturnValue);
            return Nota;
        }
        public string Bearings()
        {
            StringBuilder SQL = new StringBuilder();
            SQL.AppendLine("SELECT  ");
            SQL.AppendLine("	  REPLACE( ");
            SQL.AppendLine("		REPLACE( ");
            SQL.AppendLine("		STUFF(( ");
            SQL.AppendLine("				SELECT  ");
            SQL.AppendLine("					 FE.NAME ");
            SQL.AppendLine("				FROM  ");
            SQL.AppendLine("					TREEELEM TE ");
            SQL.AppendLine("					, POINT PT ");
            SQL.AppendLine("					, FREQASSIGN FA ");
            SQL.AppendLine("					, FREQENTRIES FE ");
            SQL.AppendLine("				WHERE ");
            SQL.AppendLine("					TE.TREEELEMID=PT.ELEMENTID ");
            SQL.AppendLine("					AND CAST(PT.VALUESTRING as  nvarchar)=CAST((select REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASDD_MicrologDAD') as nvarchar) ");
            SQL.AppendLine("                    AND FE.GEOMETRYTABLE='SKFCM_ASGI_BearingFreqType'");
            SQL.AppendLine("					AND fa.ELEMENTID=PT.ELEMENTID ");
            SQL.AppendLine("					AND Fa.FSID=FE.FSID ");
            SQL.AppendLine("					AND PT.ELEMENTID in (SELECT TREEELEMID FROM TREEELEM WHERE PARENTID=TEX.TREEELEMID) ");
            SQL.AppendLine("				 FOR XML PATH ('') ");
            SQL.AppendLine("			 ) ");
            SQL.AppendLine("			 , 1, 0, '' ");
            SQL.AppendLine("			) ");
            SQL.AppendLine("		,'<NAME>','') ");
            SQL.AppendLine("		,'</NAME>',', ') AS Dados_Tecnicos ");
            SQL.AppendLine(" FROM  ");
            SQL.AppendLine("	TREEELEM TEX ");
            SQL.AppendLine(" WHERE ");
            SQL.AppendLine("	TEX.TREEELEMID=" + TreeElemId);

            string Return = Connection.SQLtoString(SQL.ToString());
            if (Return.Length > 0)
            {
                Return = Return.Substring(0, Return.Length - 2);
            }

            return Return;

            //return null;
        }

        public SKF.RS.STB.Analyst.AlarmFlags AlarmFlag(Techniques Tec, string DataDtg = null)
        {
            AlarmFlags _return = AlarmFlags.None;
            string[] DadType = { };

            if (TreeElem != null)
            {
                try
                {
                    DadType = Tech.DadTypes(TreeElem.Connection, Tec);

                    StringBuilder SQL = new StringBuilder();
                    SQL.Append(" SELECT        ");
                    SQL.Append(" 	TREEELEM.ALARMFLAGS ");
                    //SQL.Append(" 	TOP 1 TREEELEM.ALARMFLAGS ");
                    SQL.Append(" FROM             ");
                    SQL.Append(" TREEELEM INNER JOIN ");
                    SQL.Append(" POINT ON TREEELEM.TREEELEMID = POINT.ELEMENTID INNER JOIN ");
                    SQL.Append(" MEASUREMENT ON POINT.ELEMENTID = MEASUREMENT.POINTID ");
                    SQL.Append(" WHERE        ");
                    SQL.Append("    (TREEELEM.PARENTID=" + TreeElem.TreeElemId.ToString() + ")  ");
                    SQL.Append(" 	AND (TREEELEM.ELEMENTENABLE = 1)  ");
                    SQL.Append(" 	AND (POINT.VALUESTRING IN (  ");
                    foreach (string _dad in DadType)
                    {
                        SQL.Append("'" + _dad + "',");
                    }
                    SQL = new StringBuilder().Append(SQL.ToString().Substring(0, SQL.ToString().Length - 1) + "))");
                    if (Tec == Techniques.Sensitive)
                    {
                        SQL.Append(" 	AND TREEELEM.NAME NOT LIKE '%MCD'");
                    }
                    SQL.Append(" 	AND (POINT.FIELDID IN (" + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor") + "," + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Dad_Id") + ")) ");
                    SQL.Append((DataDtg == null ? "" : " AND MEASUREMENT.DATADTG >= '" + DataDtg.Substring(0, 8) + "000000' ")); //AND '" + DataDtg.Substring(0, 8) + "235959'"));
                    SQL.Append(" ORDER BY  TREEELEM.ALARMFLAGS DESC, MEASUREMENT.DATADTG DESC  ");
                    //SQL.Append(" OPTION (HASH JOIN)  ");

                    #region SQL Antigo
                    //SQL.Append(" SELECT TOP 1 ");
                    //SQL.Append(" 	TE.ALARMFLAGS ");
                    //SQL.Append(" FROM  ");
                    //SQL.Append(" 	" + TreeElem.Connection.Owner + "TREEELEM TE ");
                    //SQL.Append(" 	, " + TreeElem.Connection.Owner + "POINT PT ");
                    //SQL.Append(" 	, " + TreeElem.Connection.Owner + "MEASUREMENT ME ");
                    //SQL.Append(" WHERE  ");
                    //SQL.Append(" 	TE.PARENTID=" + TreeElem.TreeElemId.ToString());
                    //SQL.Append(" 	AND TE.ELEMENTENABLE=1 ");
                    //SQL.Append(" 	AND PT.ELEMENTID = TE.TREEELEMID ");
                    //SQL.Append(" 	AND PT.VALUESTRING in (");
                    //foreach(string _dad in DadType)
                    //{
                    //    SQL.Append("'" + _dad  + "',");
                    //}
                    //SQL = new StringBuilder().Append( SQL.ToString().Substring(0,SQL.ToString().Length-1) + ")" );

                    //if (Tec == Techniques.Sensitive)
                    //{
                    //    SQL.Append("  AND TE.NAME NOT LIKE '%MCD'");
                    //}
                    //SQL.Append(" 	AND PT.FIELDID IN (" + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor") + "," + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Dad_Id") + ")");
                    //SQL.Append(" 	AND ME.POINTID=PT.ELEMENTID ");
                    //SQL.Append((DataDtg == null ? "" : " AND ME.DATADTG >= '" + DataDtg.Substring(0, 8) + "000000' ")); //AND '" + DataDtg.Substring(0, 8) + "235959'"));
                    //SQL.Append(" ORDER BY TE.ALARMFLAGS DESC, ME.DATADTG DESC ");
                    #endregion

                    DataTable dt_meas = TreeElem.Connection.DataTable(SQL.ToString());
                    if (dt_meas.Rows.Count > 0)
                        _return = (SKF.RS.STB.Analyst.AlarmFlags)uint.Parse(dt_meas.Rows[0][0].ToString());

                    _return = (SKF.RS.STB.Analyst.AlarmFlags)Convert.ToUInt32(TreeElem.Connection.SQLtoUInt(SQL.ToString()));

                }
                catch (Exception ex)
                {
                    _return = AlarmFlags.None;
                    GenericTools.DebugMsg("AlarmFlag() error: " + ex.Message);
                }
            }
            return _return;
        }

        public SKF.RS.STB.Analyst.AlarmFlags ScalarAlarmFlag(Techniques Tec, string DataDtg = null)
        {
            AlarmFlags _return = AlarmFlags.None;
            string[] DadType = { };

            if (TreeElem != null)
            {
                try
                {
                    DadType = Tech.DadTypes(TreeElem.Connection, Tec);

                    StringBuilder SQL = new StringBuilder();
                    SQL.AppendLine("SELECT ");
                    SQL.AppendLine("    ma.ALARMLEVEL ");
                    SQL.AppendLine("FROM ");
                    SQL.AppendLine("    TREEELEM TE ");
                    SQL.AppendLine("INNER JOIN ");
                    SQL.AppendLine("    POINT PT ON PT.ELEMENTID=TE.TREEELEMID");
                    SQL.AppendLine("INNER JOIN ");
                    SQL.AppendLine("    MEASUREMENT MS ON PT.ELEMENTID=MS.POINTID");
                    SQL.AppendLine("INNER JOIN ");
                    SQL.AppendLine("    MEASALARM MA ON MS.MEASID=MA.MEASID");
                    SQL.AppendLine("WHERE ");
                    SQL.AppendLine("    (TE.PARENTID=" + TreeElem.TreeElemId.ToString() + ")");
                    SQL.AppendLine("    AND (TE.ELEMENTENABLE = 1)");
                    SQL.AppendLine("    AND PT.VALUESTRING IN ( ");
                    foreach (string _dad in DadType)
                    {
                        SQL.Append("'" + _dad + "',");
                    }
                    SQL = new StringBuilder().AppendLine(SQL.ToString().Substring(0, SQL.ToString().Length - 1));
                    SQL.Append(")");
                    SQL.AppendLine("    AND PT.FIELDID IN (" + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor") + "," + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Dad_Id") + ")");
                    SQL.AppendLine((DataDtg == null ? "" : " AND MS.DATADTG like '" + DataDtg.Substring(0, 8) + "%' "));
                    SQL.AppendLine("    AND MA.TABLENAME='ScalarAlarm'");
                    SQL.AppendLine("ORDER BY ");
                    SQL.AppendLine("    MA.ALARMLEVEL DESC, MS.DATADTG DESC");

                    uint Alarm = TreeElem.Connection.SQLtoUInt(SQL.ToString());

                    _return = (SKF.RS.STB.Analyst.AlarmFlags)Alarm;
                }
                catch (Exception ex)
                {
                    _return = AlarmFlags.None;
                    GenericTools.DebugMsg("AlarmFlag() error: " + ex.Message);
                }
            }

            return _return;
        }

        public uint AlarmFlag_Point(Techniques Tec, string DataDtg = null)
        {
            uint _return = 0;
            string[] DadType = { };

            if (TreeElem != null)
            {
                try
                {
                    DadType = Tech.DadTypes(TreeElem.Connection, Tec);


                    StringBuilder SQL = new StringBuilder();
                    SQL.Append(" SELECT        ");
                    SQL.Append(" 	TREEELEM.TREEELEMID ");
                    SQL.Append(" FROM             ");
                    SQL.Append(" TREEELEM INNER JOIN ");
                    SQL.Append(" POINT ON TREEELEM.TREEELEMID = POINT.ELEMENTID INNER JOIN ");
                    SQL.Append(" MEASUREMENT ON POINT.ELEMENTID = MEASUREMENT.POINTID ");
                    SQL.Append(" WHERE        ");
                    SQL.Append("    (TREEELEM.PARENTID=" + TreeElem.TreeElemId.ToString() + ")  ");
                    SQL.Append(" 	AND (TREEELEM.ELEMENTENABLE = 1)  ");
                    SQL.Append(" 	AND (POINT.VALUESTRING IN (  ");
                    foreach (string _dad in DadType)
                    {
                        SQL.Append("'" + _dad + "',");
                    }
                    SQL = new StringBuilder().Append(SQL.ToString().Substring(0, SQL.ToString().Length - 1) + "))");
                    if (Tec == Techniques.Sensitive)
                    {
                        SQL.Append(" 	AND TREEELEM.NAME NOT LIKE '%MCD'");
                    }
                    SQL.Append(" 	AND (POINT.FIELDID IN (" + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor") + "," + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Dad_Id") + ")) ");
                    SQL.Append((DataDtg == null ? "" : " AND MEASUREMENT.DATADTG >= '" + DataDtg.Substring(0, 8) + "000000' ")); //AND '" + DataDtg.Substring(0, 8) + "235959'"));
                    SQL.Append(" ORDER BY  TREEELEM.ALARMFLAGS DESC, MEASUREMENT.DATADTG DESC  ");



                    DataTable dt_meas = TreeElem.Connection.DataTable(SQL.ToString());
                    if (dt_meas.Rows.Count > 0)
                        _return = uint.Parse(dt_meas.Rows[0][0].ToString());

                    #region SQL Antigo
                    //SQL.Append(" SELECT TOP 1 ");
                    //SQL.Append(" 	TE.ALARMFLAGS ");
                    //SQL.Append(" FROM  ");
                    //SQL.Append(" 	" + TreeElem.Connection.Owner + "TREEELEM TE ");
                    //SQL.Append(" 	, " + TreeElem.Connection.Owner + "POINT PT ");
                    //SQL.Append(" 	, " + TreeElem.Connection.Owner + "MEASUREMENT ME ");
                    //SQL.Append(" WHERE  ");
                    //SQL.Append(" 	TE.PARENTID=" + TreeElem.TreeElemId.ToString());
                    //SQL.Append(" 	AND TE.ELEMENTENABLE=1 ");
                    //SQL.Append(" 	AND PT.ELEMENTID = TE.TREEELEMID ");
                    //SQL.Append(" 	AND PT.VALUESTRING in (");
                    //foreach(string _dad in DadType)
                    //{
                    //    SQL.Append("'" + _dad  + "',");
                    //}
                    //SQL = new StringBuilder().Append( SQL.ToString().Substring(0,SQL.ToString().Length-1) + ")" );

                    //if (Tec == Techniques.Sensitive)
                    //{
                    //    SQL.Append("  AND TE.NAME NOT LIKE '%MCD'");
                    //}
                    //SQL.Append(" 	AND PT.FIELDID IN (" + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor") + "," + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Dad_Id") + ")");
                    //SQL.Append(" 	AND ME.POINTID=PT.ELEMENTID ");
                    //SQL.Append((DataDtg == null ? "" : " AND ME.DATADTG >= '" + DataDtg.Substring(0, 8) + "000000' ")); //AND '" + DataDtg.Substring(0, 8) + "235959'"));
                    //SQL.Append(" ORDER BY TE.ALARMFLAGS DESC, ME.DATADTG DESC ");
                    #endregion



                }
                catch (Exception ex)
                {
                    _return = uint.MinValue;
                    GenericTools.DebugMsg("AlarmFlag() error: " + ex.Message);
                }
            }
            return _return;
        }

        #region AlarmFlag Deprecated
        //public SKF.RS.STB.Analyst.AlarmFlags AlarmFlag_Vibration(string DataDtg = null)
        //{

        //    SKF.RS.STB.Analyst.AlarmFlags ReturnFlag = (SKF.RS.STB.Analyst.AlarmFlags.None);

        //        if (TreeElem != null)
        //        {

        //            uint DadType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASDD_MarlinDAD");
        //            uint SensorType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor");
        //            uint IMxType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASDD_ImxDAD");


        //            uint NonCollected = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASMD_NonCollection");
        //            uint MachineNotOperating = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASNC_MachineNotOperating");
        //            uint AutoCollectedMachineOk = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASNC_AutoCollectedMachineOk");
        //            uint Inspection = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection");

        //            DataTable TableNotes = TreeElem.Connection.DataTable("AlarmFlags", "TreeElem, Measurement",
        //                "(TreeElemId in (select TreeElemId from " +
        //                TreeElem.Connection.Owner + "TreeElem where ElementEnable=1 and ParentId=" + TreeElem.TreeElemId.ToString() + ")) " +
        //                " AND TreeElemId NOT IN (select TreeElemId from " + Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //                " and ElementEnable=1 AND TreeElemId in ( " +
        //                " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //                "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //                " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1 AND PARENTID!=2147000000))) "
        //                + (DataDtg == null ? "" : " AND DATADTG > '" +  DataDtg + "' ") + " AND POINTID=TREEELEMID ORDER BY AlarmFlags DESC, DATADTG DESC");

        //            if (TableNotes.Rows.Count > 0)
        //                ReturnFlag = (SKF.RS.STB.Analyst.AlarmFlags)Convert.ToUInt32(TableNotes.Rows[0][0].ToString());


        //            TableNotes.Dispose();
        //        }
        //        return ReturnFlag;
        //}
        //public SKF.RS.STB.Analyst.AlarmFlags AlarmFlag_Inspection(string DataDtg = null)
        //{

        //    SKF.RS.STB.Analyst.AlarmFlags ReturnFlag = (SKF.RS.STB.Analyst.AlarmFlags.None);

        //        if (TreeElem != null)
        //        {

        //            uint DadType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASDD_MarlinDAD");
        //            uint SensorType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor");

        //            uint NonCollected = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASMD_NonCollection");
        //            uint MachineNotOperating = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASNC_MachineNotOperating");
        //            uint AutoCollectedMachineOk = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASNC_AutoCollectedMachineOk");
        //            uint Inspection = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection");

        //            DataTable TableNotes = TreeElem.Connection.DataTable("AlarmFlags", "TreeElem, Measurement",
        //                "(TreeElemId in (select TreeElemId from " +
        //                TreeElem.Connection.Owner + "TreeElem where ElementEnable=1 and ParentId=" + TreeElem.TreeElemId.ToString() + ")) " +
        //                " AND TreeElemId IN (select TreeElemId from " + Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //                " and ElementEnable=1 AND TreeElemId in ( " +
        //                " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //                "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //                " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1 AND PARENTID!=2147000000))) "
        //                + (DataDtg == null ? "" : " AND DATADTG > '" + DataDtg + "' ") + " AND POINTID=TREEELEMID ORDER BY AlarmFlags DESC, DATADTG DESC");

        //            if (TableNotes.Rows.Count > 0)
        //                ReturnFlag = (SKF.RS.STB.Analyst.AlarmFlags)Convert.ToUInt32(TableNotes.Rows[0][0].ToString());


        //            TableNotes.Dispose();
        //        }
        //        return ReturnFlag;

        //}
        //public SKF.RS.STB.Analyst.AlarmFlags AlarmFlag_Derived(string DataDtg = null)
        //{

        //    SKF.RS.STB.Analyst.AlarmFlags ReturnFlag = (SKF.RS.STB.Analyst.AlarmFlags.None);

        //    if (TreeElem != null)
        //    {

        //        uint DadType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id");
        //        uint Derived = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint");

        //        DataTable TableNotes = TreeElem.Connection.DataTable("SELECT ALARMFLAGS FROM TREEELEM WHERE TREEELEMID IN (SELECT pt.ELEMENTID FROM " +
        //                    TreeElem.Connection.Owner + "POINT PT WHERE PT.ELEMENTID IN " +
        //                    "(SELECT TREEELEMID FROM TREEELEM WHERE PARENTID=" + TreeElem.TreeElemId.ToString() + " AND CONTAINERTYPE=4  and ElementEnable=1 AND PARENTID!=2147000000) " +
        //                    "AND VALUESTRING ='" + Derived + "'  and  FIELDID='" + DadType + "') ORDER BY ALARMFLAGS DESC");

        //        if (TableNotes.Rows.Count > 0)
        //            ReturnFlag = (SKF.RS.STB.Analyst.AlarmFlags)Convert.ToUInt32(TableNotes.Rows[0][0].ToString());


        //        TableNotes.Dispose();
        //    }
        //    return ReturnFlag;
        //}

        #endregion

        public Measurement LastMeas
        {
            get
            {
                GenericTools.DebugMsg("LastMeas(): Starting...");
                Measurement ReturnValue = null;

                if (TreeElem != null)
                {
                    uint LastMeasId;
                    string LastDataDtg;

                    try
                    {
                        LastDataDtg = TreeElem.Connection.SQLtoString("max(DataDtg)", "Measurement", "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() + ")");
                        LastMeasId = TreeElem.Connection.SQLtoUInt("max(MeasId)", "Measurement", "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() + ") and DataDtg='" + LastDataDtg + "'");
                        ReturnValue = new Measurement(TreeElem.Connection, LastMeasId);
                    }
                    catch (Exception ex)
                    {
                        ReturnValue = new Measurement();
                        GenericTools.DebugMsg("LastMeas() error: " + ex.Message);
                    }
                }

                GenericTools.DebugMsg("LastMeas(): " + ReturnValue.ToString());
                return ReturnValue;
            }

        }
        public Measurement LeastMeas(Techniques Tec)
        {
            Measurement _return = null;
            string[] DadType = { };
            uint LastMeasId;
            if (TreeElem != null)
            {
                try
                {
                    DadType = Tech.DadTypes(TreeElem.Connection, Tec);


                    StringBuilder SQL = new StringBuilder();

                    SQL.Append(" SELECT        ");
                    SQL.Append(" 	 MEASUREMENT.MEASID ");
                    SQL.Append(" FROM             ");
                    SQL.Append(" TREEELEM INNER JOIN ");
                    SQL.Append(" POINT ON TREEELEM.TREEELEMID = POINT.ELEMENTID INNER JOIN ");
                    SQL.Append(" MEASUREMENT ON POINT.ELEMENTID = MEASUREMENT.POINTID ");
                    SQL.Append(" WHERE        ");
                    SQL.Append("    (TREEELEM.PARENTID=" + TreeElem.TreeElemId.ToString() + ")  ");
                    SQL.Append(" 	AND (TREEELEM.ELEMENTENABLE = 1)  ");
                    SQL.Append(" 	AND (POINT.VALUESTRING IN (  ");
                    foreach (string _dad in DadType)
                    {
                        SQL.Append("'" + _dad + "',");
                    }
                    SQL = new StringBuilder().Append(SQL.ToString().Substring(0, SQL.ToString().Length - 1) + "))");
                    if (Tec == Techniques.Sensitive)
                    {
                        SQL.Append(" 	AND TREEELEM.NAME NOT LIKE '%MCD'");
                    }
                    SQL.Append(" 	AND (POINT.FIELDID IN (" + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor") + "," + Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Dad_Id") + ")) ");
                    SQL.Append(" ORDER BY MEASUREMENT.DATADTG DESC ");


                    DataTable dt_meas = TreeElem.Connection.DataTable(SQL.ToString());
                    if (dt_meas.Rows.Count > 0)
                        LastMeasId = uint.Parse(dt_meas.Rows[0][0].ToString());
                    else
                        LastMeasId = 0;


                    if (LastMeasId > 0)
                        _return = new Measurement(TreeElem.Connection, LastMeasId);
                    else
                        _return = new Measurement();

                }
                catch (Exception ex)
                {
                    _return = new Measurement();
                    GenericTools.DebugMsg("LastMeas() error: " + ex.Message);
                }



            }
            return _return;
        }

        #region LeastMeas Deprecated
        //public Measurement LastMeas_Derived
        //{
        //    get
        //    {
        //        GenericTools.DebugMsg("LastMeas(): Starting...");
        //        Measurement ReturnValue = null;

        //        if (TreeElem != null)
        //        {
        //            uint LastMeasId;
        //            string LastDataDtg;

        //            try
        //            {

        //                uint DadType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Dad_Id");
        //                uint Derived = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint");



        //                LastDataDtg = TreeElem.Connection.SQLtoString(
        //                    "SELECT MAX(DATADTG) FROM " +
        //                    TreeElem.Connection.Owner + "POINT PT, " + TreeElem.Connection.Owner + "Measurement ME WHERE PT.ELEMENTID IN " +
        //                    "(SELECT TREEELEMID FROM TREEELEM WHERE PARENTID=" + TreeElem.TreeElemId.ToString() + " AND CONTAINERTYPE=4  and ElementEnable=1 AND PARENTID!=2147000000) " +
        //                    "AND VALUESTRING ='" + Derived + "'  and  FIELDID='" + DadType + "' AND ME.POINTID=PT.ELEMENTID");

        //                //LastDataDtg = TreeElem.Connection.SQLtoString("max(DataDtg)", "Measurement",
        //                //    "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //                //    " and ElementEnable=1 AND TreeElemId in ( " +
        //                //    " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //                //    "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //                //    " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1)))"
        //                //    );

        //                //LastMeasId = TreeElem.Connection.SQLtoUInt("max(MeasId)", "Measurement",
        //                //    "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //                //    " and ElementEnable=1 AND TreeElemId in ( " +
        //                //    " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //                //    "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //                //    " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1)))" +
        //                //    " and DataDtg='" + LastDataDtg + "'"
        //                //    );

        //                LastMeasId = TreeElem.Connection.SQLtoUInt("SELECT MAX(MeasId) FROM " +
        //                    TreeElem.Connection.Owner + "POINT PT, " + TreeElem.Connection.Owner + "Measurement ME WHERE PT.ELEMENTID IN " +
        //                    "(SELECT TREEELEMID FROM TREEELEM WHERE PARENTID=" + TreeElem.TreeElemId.ToString() + " AND CONTAINERTYPE=4  and ElementEnable=1 AND PARENTID!=2147000000) " +
        //                    "AND VALUESTRING ='" + Derived + "'  and  FIELDID='" + DadType + "' AND ME.POINTID=PT.ELEMENTID and DataDtg='" + LastDataDtg + "'");

        //                ReturnValue = new Measurement(TreeElem.Connection, LastMeasId);
        //            }
        //            catch (Exception ex)
        //            {
        //                ReturnValue = new Measurement();
        //                GenericTools.DebugMsg("LastMeas() error: " + ex.Message);
        //            }
        //        }

        //        GenericTools.DebugMsg("LastMeas(): " + ReturnValue.ToString());
        //        return ReturnValue;
        //    }

        //}

        //public Measurement LastMeas_Vibration
        //{
        //    get
        //    {
        //        GenericTools.DebugMsg("LastMeas(): Starting...");
        //        Measurement ReturnValue = null;

        //        if (TreeElem != null)
        //        {
        //            uint LastMeasId;
        //            string LastDataDtg;

        //            try
        //            {

        //                uint DadType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASDD_MicrologDAD");
        //                uint IMxType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASDD_ImxDAD");
        //                uint SensorType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor");


        //                LastDataDtg = TreeElem.Connection.SQLtoString("max(DataDtg)", "Measurement", 
        //                    "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //                    " and ElementEnable=1 AND TreeElemId in (" +
        //                    " (select elementid from point where VALUESTRING in ('" + DadType + "','" + IMxType + "') OR  elementid in " +
        //                    "  (select elementid from point where FIELDID=" + SensorType + " AND VALUESTRING='MCD')" +
        //                    " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1))))"
        //                    );

        //                LastMeasId = TreeElem.Connection.SQLtoUInt("max(MeasId)", "Measurement", 
        //                    "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //                    " and ElementEnable=1 AND TreeElemId in (" +
        //                    " (select elementid from point where VALUESTRING in ('" + DadType + "','" + IMxType + "') OR  elementid in " +
        //                    "  (select elementid from point where FIELDID=" + SensorType + " AND VALUESTRING='MCD')" +
        //                    " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + "  and ElementEnable=1)))" +
        //                    ") and DataDtg='" + LastDataDtg + "'"
        //                    );
        //                ReturnValue = new Measurement(TreeElem.Connection, LastMeasId);
        //            }
        //            catch (Exception ex)
        //            {
        //                ReturnValue = new Measurement();
        //                GenericTools.DebugMsg("LastMeas() error: " + ex.Message);
        //            }
        //        }

        //        GenericTools.DebugMsg("LastMeas(): " + ReturnValue.ToString());
        //        return ReturnValue;
        //    }

        //}
        //public Measurement LastMeas_Inspection
        //{
        //    get
        //    {
        //        GenericTools.DebugMsg("LastMeas(): Starting...");
        //        Measurement ReturnValue = null;

        //        if (TreeElem != null)
        //        {
        //            uint LastMeasId;
        //            string LastDataDtg;

        //            try
        //            {

        //                uint DadType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASDD_MarlinDAD");
        //                uint SensorType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor");


        //                LastDataDtg = TreeElem.Connection.SQLtoString("max(DataDtg)", "Measurement",
        //                    "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //                    " and ElementEnable=1 AND TreeElemId in ( " +
        //                    " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //                    "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //                    " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1)))"
        //                    );

        //                LastMeasId = TreeElem.Connection.SQLtoUInt("max(MeasId)", "Measurement",
        //                    "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //                    " and ElementEnable=1 AND TreeElemId in ( " +
        //                    " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //                    "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //                    " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1)))" +  
        //                    " and DataDtg='" + LastDataDtg + "'"
        //                    );

        //                if (LastMeasId > 0)
        //                    ReturnValue = new Measurement(TreeElem.Connection, LastMeasId);
        //                else
        //                    ReturnValue = new Measurement();
        //            }
        //            catch (Exception ex)
        //            {
        //                ReturnValue = new Measurement();
        //                GenericTools.DebugMsg("LastMeas() error: " + ex.Message);
        //            }
        //        }

        //        GenericTools.DebugMsg("LastMeas(): " + ReturnValue.ToString());
        //        return ReturnValue;
        //    }

        //}
        #endregion


        public Machine(AnalystConnection AnalystConnection)
        {
            Connection = AnalystConnection;
        }
        public Machine(TreeElem TreeElem)
        {
            if (TreeElem.ContainerType == ContainerType.Machine)
                _TreeElem = TreeElem;
        }
        public Machine(AnalystConnection AnalystConnection, uint TreeElemId)
        {
            GenericTools.DebugMsg("Machine(" + TreeElemId.ToString() + "): Starting...");

            try
            {
                TreeElem TreeElemTemp = new TreeElem(AnalystConnection, TreeElemId);
                if (TreeElemTemp.ContainerType == ContainerType.Machine)
                    _TreeElem = TreeElemTemp;
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Machine(" + TreeElemId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("Machine(" + TreeElemId.ToString() + "): Finished!");
        }

        public uint GetConditionIndex(Techniques Tec)
        {
            int ReturnValue = 0;

            int Weitght_Alarm = 3;
            int Weitght_Alert = 2;
            int Weitght_Normal = 1;

            int CountTotal = 0;
            int CountNormal = 0;
            int CountDanger = 0;
            int CountAlert = 0;
            int Pontuacao = 0;

            try
            {
                GenericTools.DebugMsg("GetConditionIndex, Pegando lista de pontos...");
                List<Point> pPoints = GetPointList(Tec);

                GenericTools.DebugMsg("GetConditionIndex, Lista de Pontos Ok: " + pPoints.Count + " pontos");
                if (pPoints.Count > 0)
                {
                    GenericTools.DebugMsg("GetConditionIndex, Iniciando contagens...");

                    CountTotal = pPoints.Where(ponto => ponto.TreeElem.AlarmFlags != AlarmFlags.None && ponto.TreeElem.AlarmFlags != AlarmFlags.Invalid).Count();
                    GenericTools.DebugMsg("GetConditionIndex, Contando Pontos CountTotal: " + CountTotal + " pontos");

                    CountNormal = pPoints.Where(ponto => ponto.TreeElem.AlarmFlags == AlarmFlags.Good).Count() * Weitght_Normal;
                    GenericTools.DebugMsg("GetConditionIndex, Contando Pontos CountNormal: " + CountNormal + " pontos");

                    CountDanger = pPoints.Where(ponto => ponto.TreeElem.AlarmFlags == AlarmFlags.Danger).Count() * Weitght_Alarm;
                    GenericTools.DebugMsg("GetConditionIndex, Contando Pontos CountDanger: " + CountDanger + " pontos");

                    CountAlert = pPoints.Where(ponto => ponto.TreeElem.AlarmFlags == AlarmFlags.Alert).Count() * Weitght_Alert;
                    GenericTools.DebugMsg("GetConditionIndex, Contando Pontos CountAlert: " + CountAlert + " pontos");


                    Pontuacao = (CountNormal + CountAlert + CountDanger) * 100;
                    GenericTools.DebugMsg("GetConditionIndex, Pontuação: " + Pontuacao);


                    if (Pontuacao > 0)
                        ReturnValue = (Pontuacao / CountTotal);

                    if (Pontuacao == CountTotal)
                    {
                        ReturnValue = 001;
                    }

                    pPoints.Clear();
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Error: GetConditionIndex(CountTotal: " + CountTotal +
                                                       " - CountNormal: " + CountNormal +
                                                       " - CountDanger: " + CountDanger +
                                                       " - CountAlert: " + CountAlert +
                                                       " - Pontuacao: " + Pontuacao +
                                                       " - " + ex.Message.ToString() + "): Finished!");
            }

            return Convert.ToUInt32(Math.Ceiling((double)ReturnValue));
        }

        public uint GetConditionIndex(Techniques[] Tec)
        {
            int ReturnValue = 0;

            int Weitght_Alarm = 3;
            int Weitght_Alert = 2;
            int Weitght_Normal = 1;

            int CountTotal = 0;
            int CountNormal = 0;
            int CountDanger = 0;
            int CountAlert = 0;
            int Pontuacao = 0;

            try
            {
                GenericTools.DebugMsg("GetConditionIndex, Pegando lista de pontos...");

                List<Point> pPoints = new List<Point>();

                foreach (Techniques Tech in Tec)
                {
                    pPoints = pPoints.Concat(GetPointList(Tech)).ToList();
                }

                //List<Point> pPoints = GetPointList(Tec);

                GenericTools.DebugMsg("GetConditionIndex, Lista de Pontos Ok: " + pPoints.Count + " pontos");
                if (pPoints.Count > 0)
                {
                    GenericTools.DebugMsg("GetConditionIndex, Iniciando contagens...");

                    CountTotal = pPoints.Where(ponto => ponto.TreeElem.AlarmFlags != AlarmFlags.None).Count();
                    GenericTools.DebugMsg("GetConditionIndex, Contando Pontos CountTotal: " + CountTotal + " pontos");

                    CountNormal = pPoints.Where(ponto => ponto.TreeElem.AlarmFlags == AlarmFlags.Good).Count() * Weitght_Normal;
                    GenericTools.DebugMsg("GetConditionIndex, Contando Pontos CountNormal: " + CountNormal + " pontos");

                    CountDanger = pPoints.Where(ponto => ponto.TreeElem.AlarmFlags == AlarmFlags.Danger).Count() * Weitght_Alarm;
                    GenericTools.DebugMsg("GetConditionIndex, Contando Pontos CountDanger: " + CountDanger + " pontos");

                    CountAlert = pPoints.Where(ponto => ponto.TreeElem.AlarmFlags == AlarmFlags.Alert).Count() * Weitght_Alert;
                    GenericTools.DebugMsg("GetConditionIndex, Contando Pontos CountAlert: " + CountAlert + " pontos");


                    Pontuacao = (CountNormal + CountAlert + CountDanger) * 100;
                    GenericTools.DebugMsg("GetConditionIndex, Pontuação: " + Pontuacao);


                    if (Pontuacao > 0)
                        ReturnValue = (Pontuacao / CountTotal);

                    if (Pontuacao == CountTotal)
                    {
                        ReturnValue = 001;
                    }

                    pPoints.Clear();
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Error: GetConditionIndex(CountTotal: " + CountTotal +
                                                       " - CountNormal: " + CountNormal +
                                                       " - CountDanger: " + CountDanger +
                                                       " - CountAlert: " + CountAlert +
                                                       " - Pontuacao: " + Pontuacao +
                                                       " - " + ex.Message.ToString() + "): Finished!");
            }

            return Convert.ToUInt32(Math.Ceiling((double)ReturnValue));
        }
        public DataTable GetPointIdList()
        {
            DataTable ReturnValue = null;
            DataTable DVPoints = Connection.DataTable("TE.TreeElemId", "TreeElem TE", "TE.ParentId=" + TreeElemId.ToString() + " and ContainerType=4 and ElementEnable=1");
            ReturnValue = DVPoints;

            return ReturnValue;
        }
        public List<Point> GetPointList()
        {
            List<Point> ReturnValue = null;
            DataTable Points = Connection.DataTable("TE.TreeElemId", "TreeElem TE", "TE.ParentId=" + TreeElemId.ToString() + " and ContainerType=4 and ElementEnable=1");
            foreach (DataRow dr in Points.Rows)
            {
                Point ptTemp = new Point(Connection, uint.Parse(dr[0].ToString()));
                ReturnValue.Add(ptTemp);
            }
            return ReturnValue;
        }
        public List<Point> GetPointList(Techniques Tec = Techniques.All)
        {

            GenericTools.DebugMsg(" ----- GetPointList, Conexão: " + Connection.InitialCatalog + "...");
            GenericTools.DebugMsg(" ----- GetPointList, Criando objeto para lista de pontos, tecnica selecionada: " + Tec.ToString() + "...");
            List<Point> ReturnValue = new List<Point>();

            GenericTools.DebugMsg(" ----- GetPointList, Consultando pontos da maquina...");
            DataTable Points = Connection.DataTable("TE.TreeElemId", "TreeElem TE", "TE.ParentId=" + TreeElemId.ToString() + " and ContainerType=4 and ElementEnable=1");
            foreach (DataRow dr in Points.Rows)
            {
                GenericTools.DebugMsg(" ----- GetPointList DataRow, Criando Variavel PointId...");
                GenericTools.DebugMsg(" ----- GetPointList DataRow, Tentando criar ponto: [" + dr[0].ToString() + "]");
                uint PointId = Convert.ToUInt32(dr[0].ToString());
                Point ptTemp = new Point(Connection, PointId);


                try
                {

                    GenericTools.DebugMsg(" ----- GetPointList DataRow, Ponto: " + ptTemp.TreeElemId.ToString());
                    GenericTools.DebugMsg(" ----- GetPointList DataRow, Verificando se o ponto é da tecnica...");

                    string DadType = Connection.SQLtoString("SELECT SIGNATURE FROM REGISTRATION WHERE REGISTRATIONID=(SELECT VALUESTRING FROM POINT WHERE ELEMENTID=" + ptTemp.TreeElemId.ToString() + " AND FIELDID=(SELECT REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASPF_Dad_Id'))");
                    GenericTools.DebugMsg(" ----- GetPointList DataRow, Verificando Tecnica...DADRegistration: " + DadType + " - Ponto: " + ptTemp.DadType);



                    if (Tec == Techniques.Vibration &&
                            (ptTemp.DadType == Dad.Microlog
                            || ptTemp.DadType == Dad.DerivedPoint
                            || ptTemp.DadType == Dad.IMx
                            || ptTemp.DadType == Dad.WMx
                            )
                        )
                    {
                        GenericTools.DebugMsg(" ----- GetPointList DataRow, Tecnica encontrada VIBRAÇÃO, adicionando ponto...");
                        ReturnValue.Add(ptTemp);

                    }
                    else if (Tec == Techniques.MCD && ptTemp.DadType == Dad.Marlin && ptTemp.PointType == PointType.MCD)
                    {
                        GenericTools.DebugMsg(" ----- GetPointList DataRow, Tecnica encontrada MCD, adicionando ponto...");
                        ReturnValue.Add(ptTemp);
                    }
                    else if (Tec == Techniques.Sensitive && ptTemp.DadType == Dad.Marlin && ptTemp.PointType != PointType.MCD)
                    {
                        GenericTools.DebugMsg(" ----- GetPointList DataRow, Tecnica encontrada SENSITIVA, adicionando ponto...");
                        ReturnValue.Add(ptTemp);
                    }
                    else if (Tec == Techniques.TrendOil && ptTemp.DadType == Dad.OilAnalysis)
                    {
                        GenericTools.DebugMsg(" ----- GetPointList DataRow, Tecnica encontrada TrendOil, adicionando ponto...");
                        ReturnValue.Add(ptTemp);
                    }
                    else if (Tec == Techniques.All)
                    {
                        GenericTools.DebugMsg(" ----- GetPointList DataRow, Tecnica encontrada TODAS, adicionando ponto...");
                        ReturnValue.Add(ptTemp);
                    }
                }
                catch (Exception ex)
                {
                    GenericTools.DebugMsg("Error: GetPointList(TreeElemId: " + TreeElemId +
                                       " - Points Count: " + Points.Rows.Count +
                                       " - ptTemp dad: " + ptTemp.DadType.ToString() +
                                       " - ptTemp id: " + ptTemp.TreeElemId +
                                       " - Erro: " + ex.Message +
                                       "): Finished!");
                }

            }
            return ReturnValue;
        }
        public bool DerivatedCalc()
        {
            bool ReturnValue = false;

            DataTable DVPoints = Connection.DataTable("*", "TreeElem TE", "TE.ParentId=" + TreeElemId.ToString() + " and exists (select * from " + Connection.Owner + "Point Po where Po.ElementId=TE.TreeElemId and Po.FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id").ToString() + " and Po.ValueString='" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
            if (DVPoints.Rows.Count > 0)
            {
                ReturnValue = true;
                foreach (DataRow DVPointRow in DVPoints.Rows)
                    ReturnValue = ReturnValue & (new SKF.RS.STB.Analyst.Point(Connection, Convert.ToUInt32(DVPointRow["TreeElemId"]))).DerivatedCalc();
            }

            return ReturnValue;
        }
        public bool DerivatedAverageCalc()
        {
            bool ReturnValue = false;

            DataTable DVPoints = Connection.DataTable("*", "TreeElem TE", "TE.ParentId=" + TreeElemId.ToString() + " and exists (select * from " + Connection.Owner + "Point Po where Po.ElementId=TE.TreeElemId and Po.FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id").ToString() + " and Po.ValueString='" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "') and NAME like 'DVA%'");
            if (DVPoints.Rows.Count > 0)
            {
                ReturnValue = true;
                foreach (DataRow DVPointRow in DVPoints.Rows)
                    ReturnValue = ReturnValue & (new SKF.RS.STB.Analyst.Point(Connection, Convert.ToUInt32(DVPointRow["TreeElemId"]))).DerivatedAverageCalc();
            }

            return ReturnValue;
        }
        public void FixConditionalPoints()
        {
            DataTable TreeElemChilds = _TreeElem.Connection.DataTable("TreeElemId", "TreeElem", "ParentId=" + _TreeElem.TreeElemId.ToString());

            Point Point;

            if (TreeElemChilds.Rows.Count > 0)
                foreach (DataRow Row in TreeElemChilds.Rows)
                {
                    Point = new Point(_TreeElem.Connection, (uint)Row["TreeElemId"]);
                    Point.FixConditionalPoints();
                }
        }


        #region Notes

        public List<SKF.RS.STB.Analyst.Note> Notes(Techniques _tec) { return Notes(_tec, DateTime.MinValue, DateTime.MaxValue); }
        public List<SKF.RS.STB.Analyst.Note> Notes(Techniques _tec, DateTime Date)
        {
            Date = new DateTime(Date.Year, Date.Month, Date.Day, 0, 0, 0);
            return Notes(_tec, Date, Date.AddDays(1));
        }
        public List<SKF.RS.STB.Analyst.Note> Notes(Techniques _tec, DateTime StartDateTime, DateTime EndDateTime)
        {
            List<SKF.RS.STB.Analyst.Note> ReturnValue = new List<SKF.RS.STB.Analyst.Note>();

            if (TreeElem != null)
            {
                List<Point> pPoints = GetPointList(_tec);
                string sPoints = "'" + TreeElemId.ToString() + "',";

                foreach (Point _point in pPoints)
                {
                    sPoints += "'" + _point.TreeElemId + "',";
                }
                sPoints = sPoints.Substring(0, sPoints.Length - 1);
                DataTable TableNotes = TreeElem.Connection.DataTable("NotesId", "Notes",
                            " OwnerId in (" + sPoints + ") " +
                            " and (DataDtg between '" + GenericTools.DateTime(StartDateTime) + "' and '" + GenericTools.DateTime(EndDateTime) + "') ORDER BY DataDtg DESC");


                //DataTable TableNotes = TreeElem.Connection.DataTable("NotesId", "Notes", "(OwnerId=" + TreeElem.TreeElemId.ToString() + " or OwnerId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ElementEnable=1 and ParentId=" + TreeElem.TreeElemId.ToString() + ")) and (DataDtg between '" + GenericTools.DateTime(StartDateTime) + "' and '" + GenericTools.DateTime(EndDateTime) + "')");-
                if (TableNotes.Rows.Count > 0)
                    for (int i = 0; i < TableNotes.Rows.Count; i++)
                        ReturnValue.Add(new Note(Connection, Convert.ToUInt32(TableNotes.Rows[i]["NotesId"])));
                TableNotes.Dispose();
            }
            return ReturnValue;
        }



        public List<SKF.RS.STB.Analyst.Note> Notes() { return Notes(DateTime.MinValue, DateTime.MaxValue); }
        public List<SKF.RS.STB.Analyst.Note> Notes(DateTime Date)
        {
            Date = new DateTime(Date.Year, Date.Month, Date.Day, 0, 0, 0);
            return Notes(Date, Date.AddDays(1));
        }
        public List<SKF.RS.STB.Analyst.Note> Notes(DateTime StartDateTime, DateTime EndDateTime)
        {
            List<SKF.RS.STB.Analyst.Note> ReturnValue = new List<SKF.RS.STB.Analyst.Note>();

            if (TreeElem != null)
            {
                DataTable TableNotes = TreeElem.Connection.DataTable("NotesId", "Notes", "(OwnerId=" + TreeElem.TreeElemId.ToString() + " or OwnerId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ElementEnable=1 and ParentId=" + TreeElem.TreeElemId.ToString() + ")) and (DataDtg between '" + GenericTools.DateTime(StartDateTime) + "' and '" + GenericTools.DateTime(EndDateTime) + "')");
                if (TableNotes.Rows.Count > 0)
                    for (int i = 0; i < TableNotes.Rows.Count; i++)
                        ReturnValue.Add(new Note(Connection, Convert.ToUInt32(TableNotes.Rows[i]["NotesId"])));
                TableNotes.Dispose();
            }
            return ReturnValue;
        }

        #region Notes Deprecated
        //public List<SKF.RS.STB.Analyst.Note> Notes_Inspection() { return Notes_Inspection(DateTime.MinValue, DateTime.MaxValue); }
        //public List<SKF.RS.STB.Analyst.Note> Notes_Inspection(DateTime Date)
        //{
        //    Date = new DateTime(Date.Year, Date.Month, Date.Day, 0, 0, 0);
        //    return Notes_Inspection(Date, Date.AddDays(1));
        //}
        //public List<SKF.RS.STB.Analyst.Note> Notes_Inspection(DateTime StartDateTime, DateTime EndDateTime)
        //{
        //    List<SKF.RS.STB.Analyst.Note> ReturnValue = new List<SKF.RS.STB.Analyst.Note>();
        //    if (TreeElem != null)
        //    {

        //        uint DadType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASDD_MarlinDAD");
        //        uint SensorType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor");

        //        uint NonCollected = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASMD_NonCollection");
        //        uint MachineNotOperating = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASNC_MachineNotOperating");
        //        uint AutoCollectedMachineOk = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASNC_AutoCollectedMachineOk");
        //        uint Inspection = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection");

        //        DataTable TableNotes = TreeElem.Connection.DataTable("NotesId", "Notes",
        //            "(OwnerId=" + TreeElem.TreeElemId.ToString() + " or OwnerId in (select TreeElemId from " +
        //            TreeElem.Connection.Owner + "TreeElem where ElementEnable=1 and ParentId=" + TreeElem.TreeElemId.ToString() + ")) " +
        //            " AND OWNERID IN ( select TreeElemId from " + Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //            " and ElementEnable=1 AND TreeElemId in ( " +
        //            " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //            "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //            " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1)))" +
        //            " and (DataDtg between '" + GenericTools.DateTime(StartDateTime) + "' and '" + GenericTools.DateTime(EndDateTime) + "') ");

        //        if (TableNotes.Rows.Count > 0)
        //            for (int i = 0; i < TableNotes.Rows.Count; i++)
        //                ReturnValue.Add(new SKF.RS.STB.Analyst.Note(Connection, Convert.ToUInt32(TableNotes.Rows[i]["NotesId"])));
        //        TableNotes.Dispose();
        //    }
        //    return ReturnValue;
        //}
        //public List<SKF.RS.STB.Analyst.Note> Notes_Vibration() { return Notes_Vibration(DateTime.MinValue, DateTime.MaxValue); }
        //public List<SKF.RS.STB.Analyst.Note> Notes_Vibration(DateTime Date)
        //{
        //    Date = new DateTime(Date.Year, Date.Month, Date.Day, 0, 0, 0);
        //    return Notes_Vibration(Date, Date.AddDays(1));
        //}
        //public List<SKF.RS.STB.Analyst.Note> Notes_Vibration(DateTime StartDateTime, DateTime EndDateTime)
        //{
        //    List<SKF.RS.STB.Analyst.Note> ReturnValue = new List<SKF.RS.STB.Analyst.Note>();
        //    if (TreeElem != null)
        //    {

        //        uint DadType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASDD_MarlinDAD");
        //        uint SensorType = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor");

        //        uint NonCollected = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASMD_NonCollection");
        //        uint MachineNotOperating = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASNC_MachineNotOperating");
        //        uint AutoCollectedMachineOk = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASNC_AutoCollectedMachineOk");
        //        uint Inspection = SKF.RS.STB.Analyst.Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection");

        //        DataTable TableNotes = TreeElem.Connection.DataTable("NotesId", "Notes",
        //            "(OwnerId=" + TreeElem.TreeElemId.ToString() + " or OwnerId in (select TreeElemId from " +
        //            TreeElem.Connection.Owner + "TreeElem where ElementEnable=1 and ParentId=" + TreeElem.TreeElemId.ToString() + ")) " +
        //            " AND OWNERID NOT IN ( select TreeElemId from " + Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //            " and ElementEnable=1 AND TreeElemId in ( " +
        //            " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //            "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //            " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1)))" +
        //            " and (DataDtg between '" + GenericTools.DateTime(StartDateTime) + "' and '" + GenericTools.DateTime(EndDateTime) + "') ");

        //        if (TableNotes.Rows.Count > 0)
        //            for (int i = 0; i < TableNotes.Rows.Count; i++)
        //                ReturnValue.Add(new SKF.RS.STB.Analyst.Note(Connection, Convert.ToUInt32(TableNotes.Rows[i]["NotesId"])));
        //        TableNotes.Dispose();
        //    }
        //    return ReturnValue;
        //}
        #endregion

        #endregion
        #region Priority - GROUPTBL
        private bool _PriorityLoaded = false;
        private Priority _Priority;
        public Priority Priority
        {
            get
            {
                if (!_PriorityLoaded)
                {
                    _Priority = (Analyst.Priority)Connection.SQLtoUInt("Priority", "GroupTbl", "ElementId=" + TreeElemId.ToString());
                    _PriorityLoaded = true;
                }
                return _Priority;
            }
            set
            {
                if (Connection.SQLUpdate("GroupTbl", "Priority", (uint)value, "ElementId=" + TreeElemId.ToString()) > 0)
                    _Priority = value;
            }
        }
        public string PriorityABC
        {
            get
            {
                switch (Priority)
                {
                    case Analyst.Priority.High:
                    case Analyst.Priority.Critical:
                        return "A";

                    case Analyst.Priority.Medium:
                        return "B";

                    case Analyst.Priority.Lowest:
                    case Analyst.Priority.Low:
                        return "C";

                    case Analyst.Priority.None:
                    default:
                        return string.Empty;
                }
            }
        }
        #endregion
        #region Asset Name - GROUPTBL
        private string _AssetName = null;
        public string AssetName
        {
            get
            {
                if (_AssetName == null)
                    _AssetName = Connection.SQLtoString("AssetName", "GroupTbl", "ElementId=" + TreeElemId.ToString());
                return _AssetName;
            }
            set
            {
                if (Connection.SQLUpdate("GroupTbl", "AssetName", value, "ElementId=" + TreeElemId.ToString()) > 0)
                    _AssetName = value;
                else
                    _AssetName = null;
            }
        }
        private string _SegmentName = null;
        public string SegmentName
        {
            get
            {
                if (_SegmentName == null)
                    _SegmentName = Connection.SQLtoString("SegmentName", "GroupTbl", "ElementId=" + TreeElemId.ToString());
                return _SegmentName;
            }
            set
            {
                if (Connection.SQLUpdate("GroupTbl", "SegmentName", value, "ElementId=" + TreeElemId.ToString()) > 0)
                    _SegmentName = value;
                else
                    _SegmentName = null;
            }
        }
        public string GetSegmentName(Segment segment)
        {
            string Segment;

            Segment = SegmentName;

            string[] allSegments = SegmentName.Split(';');

            if (allSegments.Length >= ((int)segment + 1))
            {
                if (allSegments[(int)segment] != "")
                {
                    Segment = allSegments[(int)segment].Trim();
                    return Segment;
                }
            }

            return Segment;
        }
        #endregion


        #region Machine Running

        public bool IsRunning(Techniques Tec) { return IsRunning(Tec, false); }
        public bool IsRunning(Techniques Tec, bool CheckEntireSet = false) { return IsRunning(Tec, DateTime.MinValue, DateTime.MaxValue, CheckEntireSet); }
        public bool IsRunning(Techniques Tec, DateTime Date, bool CheckEntireSet = false)
        {
            Date = new DateTime(Date.Year, Date.Month, Date.Day, 0, 0, 0);
            return IsRunning(Tec, Date, Date.AddDays(1), CheckEntireSet);
        }

        private bool IsRunning(Techniques Tec, DateTime StartDateTime, DateTime EndDateTime, bool CheckEntireSet = false)
        {
            GenericTools.DebugMsg("IsRunning Tec(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + "): Starting...");
            bool ReturnValue = true;

            try
            {

                uint NonCollected = Registration.RegistrationId(Connection, "SKFCM_ASMD_NonCollection");
                uint MachineNotOperating = Registration.RegistrationId(Connection, "SKFCM_ASNC_MachineNotOperating");
                uint AutoCollectedMachineOk = Registration.RegistrationId(Connection, "SKFCM_ASNC_AutoCollectedMachineOk");
                uint Inspection = Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection");
                uint ResultMeasurement;
                uint ReadType;

                Measurement pLastMeas = LeastMeas(Tec);
                DataTable pResult = Connection.DataTable("SELECT EXDWORDVAL1, ReadingType FROM MEASREADING WHERE MEASID=" + LastMeas.MeasId);

                ResultMeasurement = (pResult.Rows[0][0].ToString() == "" ? 0 : uint.Parse(pResult.Rows[0][0].ToString()));
                ReadType = (pResult.Rows[0][1].ToString() == "" ? 0 : uint.Parse(pResult.Rows[0][1].ToString()));

                if (Tec == Techniques.Sensitive)
                {
                    if (ReadType != Inspection)
                    {
                        //if (Notes_Inspection(pLastMeas.DataDTG, DateTime.Now).Count > 0)
                        if (Notes(Tec, pLastMeas.DataDTG, DateTime.Now).Count > 0)
                        {
                            string first_note = Notes(pLastMeas.DataDTG, DateTime.Now)[0].Text;
                            if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("MAQUINA NÃO OPERANDO") == true) && (ResultMeasurement != AutoCollectedMachineOk))
                                ReturnValue = false;

                            if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("COLETA AUTOMÁTICA DA MÁQUINA OK") == true) && (ResultMeasurement == AutoCollectedMachineOk))
                                ReturnValue = true;
                        }
                        else
                        {
                            if (ReadType == NonCollected && ResultMeasurement == AutoCollectedMachineOk)
                                ReturnValue = true;

                            if (ReadType == NonCollected && ResultMeasurement == MachineNotOperating)
                                ReturnValue = false;
                        }
                    }
                    else
                        ReturnValue = true;
                }

                if (Tec == Techniques.MCD)
                {
                    Point pt = new Point(TreeElem.Connection, pLastMeas.PointId);
                    if (pt.PointType == PointType.MCD)
                    {
                        List<SKF.RS.STB.Analyst.Note> Notes1 = Notes(Tec, pLastMeas.DataDTG, DateTime.Now);
                        if (Notes1.Count > 0)
                        {
                            string first_note = Notes1[0].Text;
                            DateTime Date_note = Notes1[0].DataDtg;

                            if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("MAQUINA NÃO OPERANDO") == true) && (pLastMeas.DataDTG < Date_note))
                                ReturnValue = false;
                            else
                            {
                                if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("MAQUINA NÃO OPERANDO") == true) && (ResultMeasurement != AutoCollectedMachineOk))
                                    ReturnValue = false;
                                if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("COLETA AUTOMÁTICA DA MÁQUINA OK") == true) && (ResultMeasurement == AutoCollectedMachineOk))
                                    ReturnValue = true;
                            }
                        }
                        else
                        {
                            if (ReadType == NonCollected && ResultMeasurement == AutoCollectedMachineOk)
                                ReturnValue = true;
                            if (ReadType == NonCollected && ResultMeasurement == MachineNotOperating)
                                ReturnValue = false;
                        }
                    }
                }

                if (Tec == Techniques.Vibration)
                {
                    List<SKF.RS.STB.Analyst.Note> Notes1 = Notes(Tec, pLastMeas.DataDTG, DateTime.Now);

                    if (Notes1.Count > 0)
                    {
                        string first_note = Notes1[0].Text;
                        DateTime Date_note = Notes1[0].DataDtg;

                        if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("MAQUINA NÃO OPERANDO") == true) && (pLastMeas.DataDTG < Date_note))
                            ReturnValue = false;
                        else
                            ReturnValue = IsRunning(StartDateTime, EndDateTime);
                    }
                    else
                        ReturnValue = IsRunning(StartDateTime, EndDateTime, true, Tec);
                }

            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("IsRunning Tec(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("IsRunning Tec(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }




        #region IsRunning Deprecated
        //private bool IsRunning_Inspection(DateTime StartDateTime, DateTime EndDateTime, bool CheckEntireSet = false)
        //{
        //    GenericTools.DebugMsg("IsRunning_Inspection(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + "): Starting...");
        //    bool ReturnValue = true;

        //    try
        //    {
        //        uint DadType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASDD_MarlinDAD");
        //        uint SensorType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor");

        //        uint NonCollected = Registration.RegistrationId(Connection, "SKFCM_ASMD_NonCollection");
        //        uint MachineNotOperating = Registration.RegistrationId(Connection, "SKFCM_ASNC_MachineNotOperating");
        //        uint AutoCollectedMachineOk = Registration.RegistrationId(Connection, "SKFCM_ASNC_AutoCollectedMachineOk");
        //        uint Inspection = Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection");

        //        DataTable LastDataDtg = TreeElem.Connection.DataTable(" Measid, DataDtg", "Measurement",
        //            "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //            " and ElementEnable=1 AND TreeElemId in ( " +
        //            " select elementid from point where VALUESTRING='" + DadType + "' and elementid not in ( " +
        //            "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //            " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1))) ORDER BY DataDtg DESC"
        //            );


        //        string LastMeasId = LastDataDtg.Rows[0][0].ToString();

        //        uint ResultMeasurement = Connection.SQLtoUInt("SELECT EXDWORDVAL1 FROM MEASREADING WHERE MEASID=" + LastMeasId);
        //        uint ReadingType = Connection.SQLtoUInt("SELECT ReadingType FROM MEASREADING WHERE MEASID=" + LastMeasId);

        //        if (ReadingType != Inspection) 
        //        {
        //            if (Notes_Inspection(GenericTools.StrToDateTime(LastDataDtg.Rows[0][1].ToString()), DateTime.Now).Count > 0)
        //            {
        //                string first_note = Notes(GenericTools.StrToDateTime(LastDataDtg.Rows[0][1].ToString()), DateTime.Now)[0].Text;
        //                if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("MAQUINA NÃO OPERANDO") == true) && (ResultMeasurement != AutoCollectedMachineOk))
        //                {
        //                    ReturnValue = false;

        //                }
        //                if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("COLETA AUTOMÁTICA DA MÁQUINA OK") == true) && (ResultMeasurement == AutoCollectedMachineOk))
        //                {
        //                    ReturnValue = true;
        //                }
        //            }
        //            else
        //            {
        //                if (ReadingType == NonCollected && ResultMeasurement == AutoCollectedMachineOk)
        //                {
        //                    ReturnValue = true;
        //                }

        //                if (ReadingType == NonCollected && ResultMeasurement == MachineNotOperating)
        //                {
        //                    ReturnValue = false;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            ReturnValue = true;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        GenericTools.DebugMsg("IsRunning_Inspection(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + ") error: " + ex.Message);
        //    }

        //    GenericTools.DebugMsg("IsRunning_Inspection(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + "): " + ReturnValue.ToString());

        //    return ReturnValue;
        //}
        //public bool IsRunning_Vibration() { return IsRunning_Vibration(false); }
        //public bool IsRunning_Vibration(bool CheckEntireSet = false) { return IsRunning_Vibration(DateTime.MinValue, DateTime.MaxValue, CheckEntireSet); }
        //public bool IsRunning_Vibration(DateTime Date, bool CheckEntireSet = false)
        //{
        //    Date = new DateTime(Date.Year, Date.Month, Date.Day, 0, 0, 0);
        //    return IsRunning_Vibration(Date, Date.AddDays(1), CheckEntireSet);
        //}
        //private bool IsRunning_Vibration(DateTime StartDateTime, DateTime EndDateTime, bool CheckEntireSet = false)
        //{
        //    GenericTools.DebugMsg("IsRunning_Vibration(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + "): Starting...");
        //    bool ReturnValue = true;

        //    try
        //    {

        //        uint DadType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASDD_MarlinDAD");
        //        uint SensorType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASPF_Sensor");
        //       // uint IMxType = Registration.RegistrationId(TreeElem.Connection, "SKFCM_ASDD_ImxDAD");

        //        uint NonCollected = Registration.RegistrationId(Connection, "SKFCM_ASMD_NonCollection");
        //        uint MachineNotOperating = Registration.RegistrationId(Connection, "SKFCM_ASNC_MachineNotOperating");
        //        uint AutoCollectedMachineOk = Registration.RegistrationId(Connection, "SKFCM_ASNC_AutoCollectedMachineOk");

        //        DataTable LastDataDtg = TreeElem.Connection.DataTable(" Measid, DataDtg, PointId", "Measurement",
        //            "PointId in (select TreeElemId from " + TreeElem.Connection.Owner + "TreeElem where ParentId=" + TreeElem.TreeElemId.ToString() +
        //            " and ElementEnable=1 AND TreeElemId not in ( " +
        //            " select elementid from point where VALUESTRING in ('" + DadType + "') and elementid not in ( " +
        //            "  select elementid from point where fieldid='" + SensorType + "' and VALUESTRING='MCD')" +
        //            " and ELEMENTID in (select TREEELEMID from TREEELEM where parentid=" + TreeElem.TreeElemId.ToString() + " and ElementEnable=1))) ORDER BY DataDtg DESC"
        //            );

        //        string LastMeasId = LastDataDtg.Rows[0][0].ToString();
        //        uint PointId = Convert.ToUInt32(LastDataDtg.Rows[0][2]);

        //        Point pt = new Point(TreeElem.Connection, PointId);
        //        if (pt.PointType == PointType.MCD)
        //        {
        //            uint ResultMeasurement = Connection.SQLtoUInt("SELECT EXDWORDVAL1 FROM MEASREADING WHERE MEASID=" + LastMeasId);
        //            uint ReadingType = Connection.SQLtoUInt("SELECT ReadingType FROM MEASREADING WHERE MEASID=" + LastMeasId);

        //            List<SKF.RS.STB.Analyst.Note> Notes1 = Notes(GenericTools.StrToDateTime(LastDataDtg.Rows[0][1].ToString()), DateTime.Now);


        //            if (Notes1.Count > 0)
        //            {
        //                string first_note = Notes1[0].Text;
        //                DateTime Date_note = Notes1[0].DataDtg;

        //                if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("MAQUINA NÃO OPERANDO") == true) && (GenericTools.StrToDateTime(LastDataDtg.Rows[0][1].ToString()) < Date_note))
        //                {
        //                    ReturnValue = false;
        //                }
        //                else
        //                {
        //                    if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("MAQUINA NÃO OPERANDO") == true) && (ResultMeasurement != AutoCollectedMachineOk))
        //                    {
        //                        ReturnValue = false;

        //                    }
        //                    if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("COLETA AUTOMÁTICA DA MÁQUINA OK") == true) && (ResultMeasurement == AutoCollectedMachineOk))
        //                    {
        //                        ReturnValue = true;
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                if (ReadingType == NonCollected && ResultMeasurement == AutoCollectedMachineOk)
        //                {
        //                    ReturnValue = true;
        //                }

        //                if (ReadingType == NonCollected && ResultMeasurement == MachineNotOperating)
        //                {
        //                    ReturnValue = false;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            List<SKF.RS.STB.Analyst.Note> Notes1 = Notes_Vibration(GenericTools.StrToDateTime(LastDataDtg.Rows[0][1].ToString()), DateTime.Now);

        //            if (Notes1.Count > 0)
        //            {
        //                string first_note = Notes1[0].Text;
        //                DateTime Date_note = Notes1[0].DataDtg;

        //                if ((first_note.ToUpper().Contains("MACHINE NOT OPERATING") == true || first_note.ToUpper().Contains("MAQUINA NÃO OPERANDO") == true) && (GenericTools.StrToDateTime(LastDataDtg.Rows[0][1].ToString()) < Date_note))
        //                {
        //                    ReturnValue = false;
        //                }
        //                else
        //                {
        //                    ReturnValue = IsRunning(StartDateTime, EndDateTime);
        //                }
        //            }
        //            else
        //            {
        //                ReturnValue = IsRunning(StartDateTime, EndDateTime);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        GenericTools.DebugMsg("IsRunning_Vibration(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + ") error: " + ex.Message);
        //    }

        //    GenericTools.DebugMsg("IsRunning_Vibration(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + "): " + ReturnValue.ToString());

        //    return ReturnValue;
        //}
        #endregion

        public bool IsRunning() { return IsRunning(false); }
        public bool IsRunning(bool CheckEntireSet = false) { return IsRunning(DateTime.MinValue, DateTime.MaxValue, CheckEntireSet); }
        public bool IsRunning(DateTime Date, bool CheckEntireSet = false)
        {
            Date = new DateTime(Date.Year, Date.Month, Date.Day, 0, 0, 0);
            return IsRunning(Date, Date.AddDays(1), CheckEntireSet);
        }
        private bool IsRunning(DateTime StartDateTime, DateTime EndDateTime, bool CheckEntireSet = false, Techniques Tec = Techniques.All)
        {
            GenericTools.DebugMsg("IsRunning(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + "): Starting...");
            bool ReturnValue = true;

            try
            {
                uint SitPoint = Connection.SQLtoUInt("TreeElemId", "TreeElem", "upper(Name) like 'MI SIT%' and ElementEnable=1 and ParentEnable=0 and ParentId=" + TreeElemId.ToString());
                if (SitPoint > 0)
                {
                    string LastMeasDtg = Connection.SQLtoString("max(DataDtg)", "MeasDtsRead", "PointId=" + SitPoint.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and (DataDtg between '" + GenericTools.DateTime(StartDateTime) + "' and '" + GenericTools.DateTime(EndDateTime) + "')");
                    uint LastReadingId = Connection.SQLtoUInt("max(ReadingId)", "MeasDtsRead", "PointId=" + SitPoint.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and DataDtg='" + LastMeasDtg + "'");
                    uint LastMeasOverall = Connection.SQLtoUInt("OverallValue", "MeasDtsRead", "ReadingId=" + LastReadingId.ToString());
                    ReturnValue = (!(LastMeasOverall == 2));
                }
                else if (CheckEntireSet)
                {
                    Machine TempMachine;
                    foreach (TreeElem Child in TreeElem.Parent.Child)
                        if (Child.TreeElemId != TreeElemId)
                        {
                            TempMachine = new Machine(Child);
                            ReturnValue = ReturnValue & TempMachine.IsRunning(StartDateTime, EndDateTime, false, Tec);
                        }
                }
                foreach (SKF.RS.STB.Analyst.Note NoteItem in Notes(Tec, StartDateTime, EndDateTime))
                {
                    if (NoteItem.Text.ToUpper().Contains("MACHINE NOT OPERATING")) ReturnValue = false;
                    //if (NoteItem.Category == NoteCategory.NonCollectionNote) ReturnValue = false;
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("IsRunning(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("IsRunning(" + StartDateTime.ToString() + ", " + EndDateTime.ToString() + ", " + CheckEntireSet.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }
        #endregion
        #region Slotnumber
        /// <summary>SlotNumer_Sort.</summary>
        /// <param name="COLUMN">Column to reference sort</param>
        /// <returns>It will sort points with referenced FIELD ex NAME or more than one NAME, SLOTNUMBER, ETC. </returns>
        public void SlotNumber_Sort(string field)
        {

            DataTable PointsToReply = new DataTable();
            PointsToReply = Connection.DataTable("TREEELEMID, NAME", "TREEELEM", "PARENTID=" + this.TreeElemId + " order by " + (field == null ? "NAME" : field));

            DataTable Sorted = new DataTable();
            Sorted.Columns.Add("TreeElemId", typeof(int));

            foreach (DataRow dr in PointsToReply.Rows)
            {
                if (dr["name"].ToString().Trim() == "MI SIT")
                {
                    DataRow dr_in = Sorted.NewRow();
                    dr_in["TreeElemId"] = dr["TREEELEMID"].ToString();
                    Sorted.Rows.Add(dr_in);
                    break;
                }
            }
            foreach (DataRow dr in PointsToReply.Rows)
            {
                if (dr["name"].ToString().StartsWith("MI") && dr["name"].ToString().Trim() != "MI SIT")
                {
                    DataRow dr_in = Sorted.NewRow();
                    dr_in["TreeElemId"] = dr["TREEELEMID"].ToString();
                    Sorted.Rows.Add(dr_in);
                }
            }
            foreach (DataRow dr in PointsToReply.Rows)
            {
                if (dr["name"].ToString().StartsWith("DV"))
                {
                    DataRow dr_in = Sorted.NewRow();
                    dr_in["TreeElemId"] = dr["TREEELEMID"].ToString();
                    Sorted.Rows.Add(dr_in);
                }
            }

            if (Sorted.Rows.Count > 0)
            {
                int i = 1;
                foreach (DataRow PointsRow in Sorted.Rows)
                {
                    Connection.SQLUpdate("TREEELEM", "SLOTNUMBER", i, "TREEELEMID=" + PointsRow["TREEELEMID"].ToString());
                    i++;
                }
            }
        }
        /// <summary>SlotNumer_Sort.</summary>
        /// <param name="COLUMN">Column to reference sort</param>
        /// <returns>It will sort points with referenced FIELD ex NAME or more than one NAME, SLOTNUMBER, ETC. </returns>
        public void SlotNumber_Sort(AnalystConnection Connection, int parentid, string field)
        {

            DataTable PointsToReply = new DataTable();
            PointsToReply = Connection.DataTable("TREEELEMID", "TREEELEM", "PARENTID=" + parentid + " order by " + (field == null ? "NAME" : field));

            if (PointsToReply.Rows.Count > 0)
            {
                int i = 1;
                foreach (DataRow PointsRow in PointsToReply.Rows)
                {
                    Connection.SQLUpdate("TREEELEM", "SLOTNUMBER", i, "TREEELEMID=" + PointsRow["TREEELEMID"].ToString());
                    i++;
                }
            }
        }
        #endregion
        #region Situation Point
        private uint _SITPoint = 0;
        public uint SITPoint
        {
            get
            {
                if (_SITPoint == 0)
                    _SITPoint = Connection.SQLtoUInt("TreeElemId", "TreeElem", "upper(Name) like 'MI SIT%' and ElementEnable=1 and ParentEnable=0 and ParentId=" + TreeElemId.ToString());
                return _SITPoint;
            }
        }
        public uint SITCreate()
        {
            uint NewTreeElemId = 0;
            if (this.SITPoint == 0)
            {
                GenericTools.DebugMsg("SITCreate(" + this.TreeElemId + "): Starting");

                DataTable TreeElemTable;
                if (Connection.IsConnected)
                {

                    uint PointId = Connection.SQLtoUInt("TREEELEMID", "TREEELEM", "PARENTID=" + this.TreeElemId);
                    TreeElem TreeElemP = new TreeElem(this.Connection, PointId);

                    switch (Connection.DBType)
                    {
                        #region Conexão Oracle
                        case DBType.Oracle: //Oracle
                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values ("
                                + "TreeElemId_Seq.NextVal, " //TreeElemId
                                + TreeElemP.HierarchyId.ToString() + ", " //HierarchyId
                                + TreeElemP.BranchLevel.ToString() + ", " //BranchLevel
                                + "-1, " //SlotNumber
                                + TreeElemP.TblSetId.ToString() + ", " //TblSetId
                                + "'MI SIT', " //Name
                                + "4, " //ContainerType
                                + "'@ 1-Operacao;2-Parado;3-Manutenc', " //Description
                                + (TreeElemP.ElementEnable ? "1, " : "0, ") //ElementEnable
                                + (TreeElemP.ParentEnable ? "1, " : "0, ") //ParentEnable
                                + "1, " //HierarchyType
                                + "1, " //AlarmFlags
                                + TreeElemP.ParentId.ToString() + ", " //ParentId
                                + TreeElemP.ParentId.ToString() + ", " //ParentRefId
                                + "0, " //ReferenceId
                                + "0, " //Good
                                + "0, " //Alert
                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                );

                            //NewTreeElemId = Connection.SQLtoUInt(Connection.Owner + "TreeElemId_Seq.CurrVal", "Dual");
                            NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_MicrologDAD") + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_Wildcard") + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '3')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 2, '0')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 3, 'Manual')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");

                            TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + "and ParentId!=2147000000");
                            if (TreeElemTable.Rows.Count > 0)
                            {
                                for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                {


                                    Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                        + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                        + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                        + "-1, " //SlotNumber
                                        + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                        + "'MI SIT', " //Name
                                        + "4, " //ContainerType
                                        + "'@ 1-Operacao;2-Parado;3-Manutenc', " //Description
                                        + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                        + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                        + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                        + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                        + NewTreeElemId.ToString() + ", " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        );

                                    Log.log("Referenced by HierarchyType:" + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " and ParentId: " + TreeElemTable.Rows[i]["ParentId"].ToString());
                                    Machine MacX = new Machine(Connection, Convert.ToUInt32(TreeElemTable.Rows[i]["ParentId"]));
                                    MacX.SlotNumber_Sort("SLOTNUMBER");
                                }
                            }
                            break;
                        #endregion

                        #region Conexão MSSQL
                        case DBType.MSSQL:
                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                + TreeElemP.HierarchyId.ToString() + ", " //HierarchyId
                                + TreeElemP.BranchLevel.ToString() + ", " //BranchLevel
                                + "-1, " //SlotNumber
                                + TreeElemP.TblSetId.ToString() + ", " //TblSetId
                                + "'MI SIT', " //Name
                                + " 4, " //ContainerType
                                + "'@ 1-Operacao;2-Parado;3-Manutenc', " //Description
                                + (TreeElemP.ElementEnable ? "1, " : "0, ") //ElementEnable
                                + (TreeElemP.ParentEnable ? "1, " : "0, ") //ParentEnable
                                + "1, " //HierarchyType
                                + "1, " //AlarmFlags
                                + TreeElemP.ParentId.ToString() + ", " //ParentId
                                + TreeElemP.ParentId.ToString() + ", " //ParentRefId
                                + "0, " //ReferenceId
                                + "0, " //Good
                                + "0, " //Alert
                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                );
                            NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");

                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_MicrologDAD") + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_Wildcard") + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '3')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 2, '0')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 3, 'Manual')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                            Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");

                            TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + "and ParentId!=2147000000");
                            if (TreeElemTable.Rows.Count > 0)
                            {
                                for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                {
                                    Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                        + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                        + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                        + "-1, " //SlotNumber
                                        + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                        + "'MI SIT', " //Name
                                        + "4, " //ContainerType
                                        + "'@ 1-Operacao;2-Parado;3-Manutenc', " //Description
                                        + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                        + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                        + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                        + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                        + NewTreeElemId.ToString() + ", " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        );

                                    Machine Mac1 = new Machine(Connection, Convert.ToUInt32(TreeElemTable.Rows[i]["ParentId"]));
                                    Mac1.SlotNumber_Sort("SLOTNUMBER");
                                }
                            }

                            break;
                        #endregion
                    }
                }
            }
            this.SlotNumber_Sort("SLOTNUMBER");
            return NewTreeElemId;
        }
        #endregion
    }

    public class Point
    {
        public bool IsLoaded { get { return (_TreeElem != null); } }

        private TreeElem _TreeElem = null;
        public TreeElem TreeElem { get { return (IsLoaded ? _TreeElem : null); } }
        public AnalystConnection Connection { get { return (IsLoaded ? _TreeElem.Connection : null); } }

        private DataTable _PointTable = null;
        public DataTable PointTable
        {
            get
            {
                if (IsLoaded && _PointTable == null && _TreeElem.Connection.IsConnected && PointId > 0)
                {
                    _PointTable = _TreeElem.Connection.DataTable("FieldId, DataType, ValueString", "Point", "ElementId=" + PointId.ToString());
                    if (_PointTable.Rows.Count < 1) _PointTable = null;
                }
                // COMENTADO POR MATEUS, POIS ELE ELIMINAVA O OBJETO QUANDO O MESMO JA EXISTIA DENTRO DO PONTO
                // DESTA FORMA QUANDO FOR CRIADO UM NOVO OBJETO ELE INICIARÁ COMO NULL
                //else
                //    _PointTable = null;

                return _PointTable;
            }
        }
        public Measurement LastMeas
        {
            get
            {
                GenericTools.DebugMsg("LastMeas(): Starting...");
                Measurement ReturnValue = null;

                if (TreeElem != null)
                {
                    uint LastMeasId;
                    string LastDataDtg;

                    try
                    {
                        LastDataDtg = TreeElem.Connection.SQLtoString("max(DataDtg)", "Measurement", "PointId=" + PointId.ToString());
                        LastMeasId = TreeElem.Connection.SQLtoUInt("max(MeasId)", "Measurement", "PointId=" + PointId.ToString() + " and DataDtg='" + LastDataDtg + "'");
                        ReturnValue = new Measurement(TreeElem.Connection, LastMeasId);
                    }
                    catch (Exception ex)
                    {
                        ReturnValue = new Measurement();
                        GenericTools.DebugMsg("LastMeas() error: " + ex.Message);
                    }
                }

                GenericTools.DebugMsg("LastMeas(): " + ReturnValue.ToString());
                return ReturnValue;
            }
        }
        public Measurement PrevLastMeas
        {
            get
            {
                GenericTools.DebugMsg("PrevLastMeas(): Starting...");
                Measurement ReturnValue = null;

                if (TreeElem != null)
                {
                    uint LastMeasId;
                    string LastDataDtg;

                    try
                    {
                        string MaxLastDataDtg = TreeElem.Connection.SQLtoString("max(DataDtg)", "Measurement", "PointId=" + PointId.ToString());

                        LastDataDtg = TreeElem.Connection.SQLtoString("max(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and DataDtg!=" + MaxLastDataDtg);
                        LastMeasId = TreeElem.Connection.SQLtoUInt("max(MeasId)", "Measurement", "PointId=" + PointId.ToString() + " and DataDtg='" + LastDataDtg + "'");
                        ReturnValue = new Measurement(TreeElem.Connection, LastMeasId);
                    }
                    catch (Exception ex)
                    {
                        ReturnValue = new Measurement();
                        GenericTools.DebugMsg("LastMeas() error: " + ex.Message);
                    }
                }

                GenericTools.DebugMsg("LastMeas(): " + ReturnValue.ToString());
                return ReturnValue;
            }
        }
        public uint PointId { get { return (TreeElem != null ? TreeElem.TreeElemId : 0); } }
        private string _PointType = string.Empty;
        public PointType PointType
        {
            get
            {
                if (string.IsNullOrEmpty(_PointType))
                {
                    uint _PointTypeId = ValueUInt("SKFCM_ASPF_Point_Type_Id");
                    if (_PointTypeId > 0)
                        _PointType = Registration.Signature(Connection, _PointTypeId);
                }

                switch (_PointType)
                {
                    case "SKFCM_ASPT_ACCurrent": return Analyst.PointType.ACCurrent;
                    case "SKFCM_ASPT_Acc": return Analyst.PointType.Acc;
                    case "SKFCM_ASPT_AccEnvelope": return Analyst.PointType.AccEnvelope;
                    case "SKFCM_ASPT_AccToDisp": return Analyst.PointType.AccToDisp;
                    case "SKFCM_ASPT_AccToVel": return Analyst.PointType.AccToVel;
                    case "SKFCM_ASPT_Battery_Level": return Analyst.PointType.Battery_Level;
                    case "SKFCM_ASPT_CountRate": return Analyst.PointType.CountRate;
                    case "SKFCM_ASPT_Counts": return Analyst.PointType.Counts;
                    case "SKFCM_ASPT_Current": return Analyst.PointType.Current;
                    case "SKFCM_ASPT_DerivedOperTimeDays": return Analyst.PointType.DerivedOperTimeDays;
                    case "SKFCM_ASPT_DerivedOperTimeHours": return Analyst.PointType.DerivedOperTimeHours;
                    case "SKFCM_ASPT_DerivedOperTimeMins": return Analyst.PointType.DerivedOperTimeMins;
                    case "SKFCM_ASPT_DerivedOperTimeMonths": return Analyst.PointType.DerivedOperTimeMonths;
                    case "SKFCM_ASPT_DerivedOperTimeSecs": return Analyst.PointType.DerivedOperTimeSecs;
                    case "SKFCM_ASPT_DerivedOperTimeWeeks": return Analyst.PointType.DerivedOperTimeWeeks;
                    case "SKFCM_ASPT_DerivedPointCalculated": return Analyst.PointType.DerivedPointCalculated;
                    case "SKFCM_ASPT_DerivedPointMARLIN": return Analyst.PointType.DerivedPointMARLIN;
                    case "SKFCM_ASPT_Displacement": return Analyst.PointType.Displacement;
                    case "SKFCM_ASPT_DisplacementMM": return Analyst.PointType.DisplacementMM;
                    case "SKFCM_ASPT_DualAcc": return Analyst.PointType.DualAcc;
                    case "SKFCM_ASPT_DualAccEnvelope": return Analyst.PointType.DualAccEnvelope;
                    case "SKFCM_ASPT_DualAccToDisp": return Analyst.PointType.DualAccToDisp;
                    case "SKFCM_ASPT_DualAccToVel": return Analyst.PointType.DualAccToVel;
                    case "SKFCM_ASPT_DualDisplacement": return Analyst.PointType.DualDisplacement;
                    case "SKFCM_ASPT_DualDisplacementMM": return Analyst.PointType.DualDisplacementMM;
                    case "SKFCM_ASPT_DualVelEnvelope": return Analyst.PointType.DualVelEnvelope;
                    case "SKFCM_ASPT_DualVelToDisp": return Analyst.PointType.DualVelToDisp;
                    case "SKFCM_ASPT_DualVelocity": return Analyst.PointType.DualVelocity;
                    case "SKFCM_ASPT_Duration": return Analyst.PointType.Duration;
                    case "SKFCM_ASPT_DynamicPressure": return Analyst.PointType.DynamicPressure;
                    case "SKFCM_ASPT_Efficiency": return Analyst.PointType.Efficiency;
                    case "SKFCM_ASPT_Flow": return Analyst.PointType.Flow;
                    case "SKFCM_ASPT_Humidity": return Analyst.PointType.Humidity;
                    case "SKFCM_ASPT_InternalTemperature": return Analyst.PointType.InternalTemperature;
                    case "SKFCM_ASPT_LinearDisplacement": return Analyst.PointType.LinearDisplacement;
                    case "SKFCM_ASPT_Logic": return Analyst.PointType.Logic;
                    case "SKFCM_ASPT_MCD": return Analyst.PointType.MCD;
                    case "SKFCM_ASPT_MicrologCyclicAcceleration": return Analyst.PointType.MicrologCyclicAcceleration;
                    case "SKFCM_ASPT_MicrologCyclicDisplacement": return Analyst.PointType.MicrologCyclicDisplacement;
                    case "SKFCM_ASPT_MicrologCyclicEnvelopeAcceleration": return Analyst.PointType.MicrologCyclicEnvelopeAcceleration;
                    case "SKFCM_ASPT_MicrologCyclicEnvelopeVelocity": return Analyst.PointType.MicrologCyclicEnvelopeVelocity;
                    case "SKFCM_ASPT_MicrologCyclicPressure": return Analyst.PointType.MicrologCyclicPressure;
                    case "SKFCM_ASPT_MicrologCyclicSEE": return Analyst.PointType.MicrologCyclicSEE;
                    case "SKFCM_ASPT_MicrologCyclicVelocity": return Analyst.PointType.MicrologCyclicVelocity;
                    case "SKFCM_ASPT_MicrologCyclicVolts": return Analyst.PointType.MicrologCyclicVolts;
                    case "SKFCM_ASPT_MicrologMotorCurrentZoom": return Analyst.PointType.MicrologMotorCurrentZoom;
                    case "SKFCM_ASPT_MicrologMotorEnvelopedCurrent": return Analyst.PointType.MicrologMotorEnvelopedCurrent;
                    case "SKFCM_ASPT_MultipleInspection": return Analyst.PointType.MultipleInspection;
                    case "SKFCM_ASPT_Noise_Level": return Analyst.PointType.Noise_Level;
                    case "SKFCM_ASPT_Oil_Analysis": return Analyst.PointType.Oil_Analysis;
                    case "SKFCM_ASPT_OrbitAcc": return Analyst.PointType.OrbitAcc;
                    case "SKFCM_ASPT_OrbitDisplacement": return Analyst.PointType.OrbitDisplacement;
                    case "SKFCM_ASPT_OrbitDisplacementMM": return Analyst.PointType.OrbitDisplacementMM;
                    case "SKFCM_ASPT_OrbitVelocity": return Analyst.PointType.OrbitVelocity;
                    case "SKFCM_ASPT_PeakHFD": return Analyst.PointType.PeakHFD;
                    case "SKFCM_ASPT_Power": return Analyst.PointType.Power;
                    case "SKFCM_ASPT_PowerFactor": return Analyst.PointType.PowerFactor;
                    case "SKFCM_ASPT_Pressure": return Analyst.PointType.Pressure;
                    case "SKFCM_ASPT_RMSHFD": return Analyst.PointType.RMSHFD;
                    case "SKFCM_ASPT_RPM": return Analyst.PointType.RPM;
                    case "SKFCM_ASPT_ResistanceMohms": return Analyst.PointType.ResistanceMohms;
                    case "SKFCM_ASPT_ResistanceOhms": return Analyst.PointType.ResistanceOhms;
                    case "SKFCM_ASPT_SPM": return Analyst.PointType.SPM;
                    case "SKFCM_ASPT_Sees": return Analyst.PointType.Sees;
                    case "SKFCM_ASPT_Signal_Level": return Analyst.PointType.Signal_Level;
                    case "SKFCM_ASPT_Signal_Quality": return Analyst.PointType.Signal_Quality;
                    case "SKFCM_ASPT_SingleInspection": return Analyst.PointType.SingleInspection;
                    case "SKFCM_ASPT_Temperature": return Analyst.PointType.Temperature;
                    case "SKFCM_ASPT_TransitionalLogic": return Analyst.PointType.TransitionalLogic;
                    case "SKFCM_ASPT_TransitionalSpeed": return Analyst.PointType.TransitionalSpeed;
                    case "SKFCM_ASPT_TriAcc": return Analyst.PointType.TriAcc;
                    case "SKFCM_ASPT_TriAccEnvelope": return Analyst.PointType.TriAccEnvelope;
                    case "SKFCM_ASPT_TriAccToDisp": return Analyst.PointType.TriAccToDisp;
                    case "SKFCM_ASPT_TriAccToVel": return Analyst.PointType.TriAccToVel;
                    case "SKFCM_ASPT_TriChannelAcc": return Analyst.PointType.TriChannelAcc;
                    case "SKFCM_ASPT_TriChannelAccEnvelope": return Analyst.PointType.TriChannelAccEnvelope;
                    case "SKFCM_ASPT_TriChannelAccToDisp": return Analyst.PointType.TriChannelAccToDisp;
                    case "SKFCM_ASPT_TriChannelAccToVel": return Analyst.PointType.TriChannelAccToVel;
                    case "SKFCM_ASPT_TriChannelDisplacement": return Analyst.PointType.TriChannelDisplacement;
                    case "SKFCM_ASPT_TriDisplacement": return Analyst.PointType.TriDisplacement;
                    case "SKFCM_ASPT_TriVelDisplacement": return Analyst.PointType.TriVelDisplacement;
                    case "SKFCM_ASPT_TriVelEnvelope": return Analyst.PointType.TriVelEnvelope;
                    case "SKFCM_ASPT_TriVelocity": return Analyst.PointType.TriVelocity;
                    case "SKFCM_ASPT_VelEnvelope": return Analyst.PointType.VelEnvelope;
                    case "SKFCM_ASPT_VelToDisp": return Analyst.PointType.VelToDisp;
                    case "SKFCM_ASPT_Velocity": return Analyst.PointType.Velocity;
                    case "SKFCM_ASPT_VoltsAc": return Analyst.PointType.VoltsAc;
                    case "SKFCM_ASPT_VoltsDc": return Analyst.PointType.VoltsDc;
                    case "SKFCM_ASPT_Wildcard": return Analyst.PointType.Wildcard;
                    default: return Analyst.PointType.None;
                }
            }
        }
        public PointTech PointTech
        {
            get
            {
                switch (PointType)
                {
                    case Analyst.PointType.Acc:
                    case Analyst.PointType.DualAcc:
                    case Analyst.PointType.MicrologCyclicAcceleration:
                    case Analyst.PointType.MicrologCyclicEnvelopeAcceleration:
                    case Analyst.PointType.TriAcc:
                    case Analyst.PointType.TriChannelAcc:
                    case Analyst.PointType.OrbitAcc:
                        return Analyst.PointTech.Accel;

                    case Analyst.PointType.AccToVel:
                    case Analyst.PointType.DualAccToVel:
                    case Analyst.PointType.DualVelocity:
                    case Analyst.PointType.MicrologCyclicVelocity:
                    case Analyst.PointType.TriAccToVel:
                    case Analyst.PointType.TriChannelAccToVel:
                    case Analyst.PointType.TriVelocity:
                    case Analyst.PointType.Velocity:
                    case Analyst.PointType.OrbitVelocity:
                        return Analyst.PointTech.Velocity;

                    case Analyst.PointType.AccEnvelope:
                    case Analyst.PointType.DualAccEnvelope:
                    case Analyst.PointType.DualVelEnvelope:
                    case Analyst.PointType.TriAccEnvelope:
                    case Analyst.PointType.TriChannelAccEnvelope:
                    case Analyst.PointType.TriVelEnvelope:
                    case Analyst.PointType.VelEnvelope:
                    case Analyst.PointType.MicrologCyclicEnvelopeVelocity:
                    case Analyst.PointType.MicrologMotorEnvelopedCurrent:
                        return Analyst.PointTech.Envelope;

                    case Analyst.PointType.AccToDisp:
                    case Analyst.PointType.Displacement:
                    case Analyst.PointType.DisplacementMM:
                    case Analyst.PointType.DualAccToDisp:
                    case Analyst.PointType.DualDisplacement:
                    case Analyst.PointType.DualDisplacementMM:
                    case Analyst.PointType.DualVelToDisp:
                    case Analyst.PointType.LinearDisplacement:
                    case Analyst.PointType.MicrologCyclicDisplacement:
                    case Analyst.PointType.OrbitDisplacement:
                    case Analyst.PointType.OrbitDisplacementMM:
                    case Analyst.PointType.TriAccToDisp:
                    case Analyst.PointType.TriChannelAccToDisp:
                    case Analyst.PointType.TriChannelDisplacement:
                    case Analyst.PointType.TriDisplacement:
                    case Analyst.PointType.TriVelDisplacement:
                    case Analyst.PointType.VelToDisp:
                        return Analyst.PointTech.Displacement;

                    case Analyst.PointType.ACCurrent:
                    case Analyst.PointType.Current:
                    case Analyst.PointType.MicrologMotorCurrentZoom:
                        return Analyst.PointTech.Current;

                    case Analyst.PointType.Flow:
                    case Analyst.PointType.Humidity:
                    case Analyst.PointType.Power:
                    case Analyst.PointType.PowerFactor:
                        return Analyst.PointTech.Process;

                    case Analyst.PointType.RMSHFD:
                    case Analyst.PointType.PeakHFD:
                        return Analyst.PointTech.HFD;

                    case Analyst.PointType.DerivedOperTimeDays:
                    case Analyst.PointType.DerivedOperTimeHours:
                    case Analyst.PointType.DerivedOperTimeMins:
                    case Analyst.PointType.DerivedOperTimeMonths:
                    case Analyst.PointType.DerivedOperTimeSecs:
                    case Analyst.PointType.DerivedOperTimeWeeks:
                    case Analyst.PointType.Duration:
                        return Analyst.PointTech.Time;

                    case Analyst.PointType.DynamicPressure:
                    case Analyst.PointType.MicrologCyclicPressure:
                    case Analyst.PointType.Pressure:
                        return Analyst.PointTech.Pressure;

                    case Analyst.PointType.MicrologCyclicVolts:
                    case Analyst.PointType.VoltsAc:
                    case Analyst.PointType.VoltsDc:
                        return Analyst.PointTech.Volts;

                    case Analyst.PointType.InternalTemperature:
                    case Analyst.PointType.Temperature:
                        return Analyst.PointTech.Temperature;

                    case Analyst.PointType.MultipleInspection:
                    case Analyst.PointType.SingleInspection:
                        return Analyst.PointTech.Inspection;

                    case Analyst.PointType.MicrologCyclicSEE:
                    case Analyst.PointType.Sees:
                        return Analyst.PointTech.SEE;

                    case Analyst.PointType.SPM:
                        return Analyst.PointTech.SPM;

                    case Analyst.PointType.Logic:
                    case Analyst.PointType.TransitionalLogic:
                        return Analyst.PointTech.Logic;

                    case Analyst.PointType.Wildcard:
                        return Analyst.PointTech.Wildcard;

                    case Analyst.PointType.Noise_Level:
                    case Analyst.PointType.Signal_Level:
                        return Analyst.PointTech.dB;

                    case Analyst.PointType.Battery_Level:
                    case Analyst.PointType.Signal_Quality:
                    case Analyst.PointType.Efficiency:
                        return Analyst.PointTech.Perc;

                    case Analyst.PointType.Oil_Analysis:
                        return Analyst.PointTech.OilAnalysis;

                    case Analyst.PointType.RPM:
                    case Analyst.PointType.TransitionalSpeed:
                        return Analyst.PointTech.Speed;

                    case Analyst.PointType.CountRate:
                    case Analyst.PointType.Counts:
                    case Analyst.PointType.DerivedPointCalculated:
                    case Analyst.PointType.DerivedPointMARLIN:
                    case Analyst.PointType.MCD:
                    case Analyst.PointType.ResistanceMohms:
                    case Analyst.PointType.ResistanceOhms:
                    default:
                        return Analyst.PointTech.None;

                }
            }
        }
        private string _Tech = string.Empty;
        public string Tech
        {
            get
            {
                if (String.IsNullOrEmpty(_Tech))
                {
                    try
                    {
                        {
                            switch (Registration.Signature(Connection, ValueUInt("SKFCM_ASPF_Point_Type_Id")))
                            {
                                case "SKFCM_ASPT_Acc":
                                    _Tech = "A";
                                    break;

                                case "SKFCM_ASPT_Velocity":
                                case "SKFCM_ASPT_AccToVel":
                                    _Tech = "V";
                                    break;

                                case "SKFCM_ASPT_AccEnvelope":
                                case "SKFCM_ASPT_VelEnvelope":
                                    _Tech = "E" + (ValueInt("SKFCM_ASPF_Input_Filter_Range") - 20599).ToString();
                                    break;

                                case "SKFCM_ASPT_Temperature":
                                    _Tech = "T";
                                    break;

                                default:
                                    _Tech = string.Empty;
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _Tech = string.Empty;
                        GenericTools.WriteLog("Point.Tech error: " + ex.Message);
                    }
                    GenericTools.DebugMsg("Point.Tech: " + _Tech);
                }

                return _Tech;
            }
        }
        private DetectionType _Detection = DetectionType.None;
        public DetectionType Detection
        {
            get
            {
                if (_Detection == DetectionType.None)
                    _Detection = (DetectionType)ValueInt("SKFCM_ASPF_Detection");
                return _Detection;
            }
            set
            {
                if (SetValueString("SKFCM_ASPF_Detection", Convert.ToUInt32(value).ToString()))
                    _Detection = value;
            }
        }

        private float _Speed = float.NaN;
        public float Speed
        {
            get
            {
                if (float.IsNaN(_Speed))
                {
                    _Speed = ValueFloat("SKFCM_ASPF_Speed");
                    if (_Speed == 0 || float.IsNaN(_Speed))
                    {
                        uint SpeedReferenceId = ValueUInt("SKFCM_ASPF_Speed_Reference_Id");
                        if (SpeedReferenceId > 0)
                            _Speed = ((new Analyst.Point(Connection, SpeedReferenceId)).LastMeas.OverallReading.OverallValue / 60) * ValueFloat("SKFCM_ASPF_Speed_Ratio");
                    }
                }
                return _Speed;
            }
            set
            {
                if (SetValueString("SKFCM_ASPF_Speed", value))
                    _Speed = value;
                else
                    _Speed = float.NaN;
            }
        }
        public float SpeedRPM { get { return (float.IsNaN(Speed) ? float.NaN : (Speed * 60)); } }

        private float _SpeedRatio = float.NaN;
        public float SpeedRatio
        {
            get
            {
                if (float.IsNaN(_SpeedRatio))
                    _SpeedRatio = ValueFloat("SKFCM_ASPF_Speed_Ratio");
                return _SpeedRatio;
            }
            set
            {
                if (SetValueString("SKFCM_ASPF_Speed_Ratio", value))
                    _SpeedRatio = value;
                else
                    _SpeedRatio = float.NaN;
            }
        }

        private float _LowFreqCutoff = float.NaN;
        public float LowFreqCutoff
        {
            get
            {
                if (float.IsNaN(_Speed))
                    _LowFreqCutoff = ValueFloat("SKFCM_ASPF_Low_Freq_Cutoff");
                return _LowFreqCutoff;
            }
            set
            {
                if (SetValueString("SKFCM_ASPF_Low_Freq_Cutoff", value))
                    _LowFreqCutoff = value;
                else
                    _LowFreqCutoff = float.NaN;
            }
        }

        private float _StartFreq = float.NaN;
        public float StartFreq
        {
            get
            {
                if (float.IsNaN(_Speed))
                    _StartFreq = ValueFloat("SKFCM_ASPF_Start_Freq");
                return _StartFreq;
            }
            set
            {
                if (SetValueString("SKFCM_ASPF_Start_Freq", value))
                    _StartFreq = value;
                else
                    _StartFreq = float.NaN;
            }
        }

        private float _EndFreq = float.NaN;
        public float EndFreq
        {
            get
            {
                if (float.IsNaN(_Speed))
                    _EndFreq = ValueFloat("SKFCM_ASPF_End_Freq");
                return _EndFreq;
            }
            set
            {
                if (SetValueString("SKFCM_ASPF_End_Freq", value))
                    _EndFreq = value;
                else
                    _EndFreq = float.NaN;
            }
        }

        public float Resolution { get { return ((float.IsNaN(StartFreq) || float.IsNaN(EndFreq) || Lines == Analyst.Lines.lines_none) ? float.NaN : ((StartFreq - EndFreq) / LinesNum)); } }

        private uint _Averages = 0;
        public uint Averages
        {
            get
            {
                if (float.IsNaN(_Speed))
                    _Averages = ValueUInt("SKFCM_ASPF_Averages");
                return _Averages;
            }
            set
            {
                if (SetValueString("SKFCM_ASPF_Averages", value))
                    _Averages = value;
            }
        }

        private Lines _Lines = Lines.lines_none;
        public Lines Lines
        {
            get
            {
                if (float.IsNaN(_Speed))
                    _Lines = (Lines)ValueFloat("SKFCM_ASPF_Lines");
                return _Lines;
            }
            set
            {
                if (SetValueString("SKFCM_ASPF_Lines", ((uint)value)))
                    _Lines = value;
            }
        }
        public float LinesNum { get { return (Convert.ToUInt32(Lines)); } }

        private float _DetectionFactor = float.NaN;
        public float DetectionFactor
        {
            get
            {
                if (float.IsNaN(_DetectionFactor))
                {
                    switch (ValueUInt("SKFCM_ASPF_Detection"))
                    {
                        case 20500:
                            _DetectionFactor = (float)1;
                            break;

                        case 20501:
                            _DetectionFactor = (float)(1 / 2);
                            break;

                        case 20502:
                            _DetectionFactor = (float)Math.Sqrt(2);
                            break;
                    }
                }
                return _DetectionFactor;
            }
        }

        public string Name { get { return TreeElem.Name; } }
        public uint TreeElemId { get { return TreeElem.TreeElemId; } }

        private ApplicationId _ApplicationId = ApplicationId.None;
        public ApplicationId ApplicationId
        {
            get
            {
                if (_ApplicationId == ApplicationId.None)
                    switch (Registration.Signature(Connection, ValueUInt("SKFCM_ASPF_Application_Id")))
                    {
                        case "SKFCM_ASAS_Cyclic":
                            _ApplicationId = Analyst.ApplicationId.Cyclic;
                            break;

                        case "SKFCM_ASAS_General":
                            _ApplicationId = Analyst.ApplicationId.General;
                            break;

                        case "SKFCM_ASAS_Inspection":
                            _ApplicationId = Analyst.ApplicationId.Inspection;
                            break;

                        case "SKFCM_ASAS_Lab_Analysis":
                            _ApplicationId = Analyst.ApplicationId.Lab_Analysis;
                            break;

                        case "SKFCM_ASAS_MotorCurrentZoom":
                            _ApplicationId = Analyst.ApplicationId.MotorCurrentZoom;
                            break;

                        case "SKFCM_ASAS_MotorEnvelopedCurrent":
                            _ApplicationId = Analyst.ApplicationId.MotorEnvelopedCurrent;
                            break;

                        case "SKFCM_ASAS_OperatingTime":
                            _ApplicationId = Analyst.ApplicationId.OperatingTime;
                            break;

                        case "SKFCM_ASAS_Orbit":
                            _ApplicationId = Analyst.ApplicationId.Orbit;
                            break;

                        case "SKFCM_ASAS_Vibration":
                            _ApplicationId = Analyst.ApplicationId.Vibration;
                            break;

                        case "SKFCM_ASAS_VibrationDualChannel":
                            _ApplicationId = Analyst.ApplicationId.VibrationDualChannel;
                            break;

                        case "SKFCM_ASAS_VibrationTriChannel":
                            _ApplicationId = Analyst.ApplicationId.VibrationTriChannel;
                            break;

                        default:
                            _ApplicationId = Analyst.ApplicationId.None;
                            break;
                    }

                return _ApplicationId;
            }
        }

        private Dad _DadType = Dad.None;
        public Dad DadType
        {
            get
            {
                if (_DadType == Dad.None)
                    //switch (Registration.Signature(Connection, ValueUInt("SKFCM_ASPF_Dad_Id")))
                    switch (Connection.SQLtoString("SELECT SIGNATURE FROM REGISTRATION WHERE REGISTRATIONID=(SELECT VALUESTRING FROM POINT WHERE ELEMENTID=" + TreeElemId.ToString() + " AND FIELDID=(SELECT REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASPF_Dad_Id'))"))
                    {
                        case "SKFCM_ASDD_DI1100DAD":
                            _DadType = Analyst.Dad.DI1100;
                            break;

                        case "SKFCM_ASDD_DerivedPoint":
                            _DadType = Analyst.Dad.DerivedPoint;
                            break;

                        case "SKFCM_ASDD_DmxDAD":
                            _DadType = Analyst.Dad.DMx;
                            break;

                        case "SKFCM_ASDD_GalDAD":
                            _DadType = Analyst.Dad.Gal;
                            break;

                        case "SKFCM_ASDD_GenDAD":
                            _DadType = Analyst.Dad.Gen;
                            break;

                        case "SKFCM_ASDD_ImxDAD":
                            _DadType = Analyst.Dad.IMx;
                            break;

                        case "SKFCM_ASDD_ImxMDAD":
                            _DadType = Analyst.Dad.IMxM;
                            break;

                        case "SKFCM_ASDD_ImxPDAD":
                            _DadType = Analyst.Dad.IMxP;
                            break;

                        case "SKFCM_ASDD_ImxSDAD":
                            _DadType = Analyst.Dad.IMxS;
                            break;

                        case "SKFCM_ASDD_ImxTDAD":
                            _DadType = Analyst.Dad.IMxT;
                            break;

                        case "SKFCM_ASDD_LmuDAD":
                            _DadType = Analyst.Dad.LMU;
                            break;

                        case "SKFCM_ASDD_ManualDAD":
                            _DadType = Analyst.Dad.Manual;
                            break;

                        case "SKFCM_ASDD_MarlinDAD":
                            _DadType = Analyst.Dad.Marlin;
                            break;

                        case "SKFCM_ASDD_MasCon_DAD":
                            _DadType = Analyst.Dad.MasCon;
                            break;

                        case "SKFCM_ASDD_Mascon16DAD":
                            _DadType = Analyst.Dad.Mascon16;
                            break;

                        case "SKFCM_ASDD_MicrologDAD":
                            _DadType = Analyst.Dad.Microlog;
                            break;

                        case "SKFCM_ASDD_MimDAD":
                            _DadType = Analyst.Dad.MIM;
                            break;

                        case "SKFCM_ASDD_OPEC_DAD":
                            _DadType = Analyst.Dad.OilAnalysis;
                            break;

                        case "SKFCM_ASDD_TmuDAD":
                            _DadType = Analyst.Dad.TMU;
                            break;

                        case "SKFCM_ASDD_WMx_Sub_WVT":
                            _DadType = Analyst.Dad.WMx_Sub_WVT;
                            break;

                        case "SKFCM_ASDD_WmxDAD":
                            _DadType = Analyst.Dad.WMx;
                            break;

                        default:
                            _DadType = Analyst.Dad.None;
                            break;
                    }

                return _DadType;
            }
        }

        public Point() { }
        public Point(TreeElem TreeElem)
        {
            if (TreeElem.ContainerType == ContainerType.Point)
                _TreeElem = TreeElem;
        }
        public Point(AnalystConnection AnalystConnection, uint TreeElemId)
        {
            GenericTools.DebugMsg("Point, Creating TreeElem element...");
            TreeElem TreeElemTemp = new TreeElem(AnalystConnection, TreeElemId);
            if (TreeElemTemp.ContainerType == ContainerType.Point)
                _TreeElem = TreeElemTemp;
        }

        private bool SetValueString(string FieldId, float NewValue) { return SetValueString(Registration.RegistrationId(_TreeElem.Connection, FieldId), NewValue); }
        private bool SetValueString(uint FieldId, float NewValue) { return SetValueString(FieldId, NewValue.ToString(System.Globalization.CultureInfo.InvariantCulture)); }
        private bool SetValueString(string FieldId, int NewValue) { return SetValueString(Registration.RegistrationId(_TreeElem.Connection, FieldId), NewValue); }
        private bool SetValueString(uint FieldId, int NewValue) { return SetValueString(FieldId, NewValue.ToString(System.Globalization.CultureInfo.InvariantCulture)); }
        private bool SetValueString(string FieldId, uint NewValue) { return SetValueString(Registration.RegistrationId(_TreeElem.Connection, FieldId), NewValue); }
        private bool SetValueString(uint FieldId, uint NewValue) { return SetValueString(FieldId, NewValue.ToString(System.Globalization.CultureInfo.InvariantCulture)); }
        private bool SetValueString(string FieldId, string NewValue) { return SetValueString(Registration.RegistrationId(_TreeElem.Connection, FieldId), NewValue); }
        private bool SetValueString(uint FieldId, string NewValue)
        {
            bool ReturnValue = false;

            ReturnValue = (Connection.SQLUpdate("Point", "ValueString", NewValue, "ElementId=" + TreeElemId.ToString() + " and FieldId=" + FieldId.ToString()) > 0);
            if (ReturnValue) _PointTable = null;

            return ReturnValue;
        }

        public string ValueString(string FieldId)
        {
            return (string)ValueString(Registration.RegistrationId(_TreeElem.Connection, FieldId));
        }
        public string ValueString(uint FieldId)
        {
            DataRow[] PointRows = PointTable.Select("FieldId=" + FieldId.ToString());
            if (PointRows.Length > 0)
                return PointRows[0]["ValueString"].ToString();
            else
                return string.Empty;
        }
        private Int32 ValueInt(string FieldId)
        {
            return Convert.ToInt32((string)ValueString(FieldId), System.Globalization.CultureInfo.InvariantCulture);
        }
        private Int32 ValueInt(uint FieldId)
        {
            return Convert.ToInt32((string)ValueString(FieldId), System.Globalization.CultureInfo.InvariantCulture);
        }
        private uint ValueUInt(string FieldId)
        {
            return Convert.ToUInt32((string)ValueString(FieldId), System.Globalization.CultureInfo.InvariantCulture);
        }
        private uint ValueUInt(uint FieldId)
        {
            return Convert.ToUInt32((string)ValueString(FieldId), System.Globalization.CultureInfo.InvariantCulture);
        }
        private float ValueFloat(string FieldId)
        {
            return Convert.ToSingle((string)ValueString(FieldId), System.Globalization.CultureInfo.InvariantCulture);
        }
        private float ValueFloat(uint FieldId)
        {
            return Convert.ToSingle((string)ValueString(FieldId), System.Globalization.CultureInfo.InvariantCulture);
        }
        /*         public uint ValueNumber(string FieldId)
        {
            return Convert.ToUInt32((string)ValueString(FieldId));
        }
        public uint ValueNumber(uint FieldId)
        {
            return Convert.ToUInt32((string)ValueString(FieldId));
        }
        */

        private AlarmMCD _AlarmMCD;
        public AlarmMCD AlarmMCD
        {
            get
            {
                if (IsLoaded)
                {
                    GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + "): Starting...");

                    AlarmMCD ReturnValue = new AlarmMCD();

                    try
                    {
                        if (Connection.IsConnected)
                        {
                            uint AlarmId = Connection.SQLtoUInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString() + " and Type=" + Registration.RegistrationId(Connection, "SKFCM_ASAT_MCD"));
                            DataTable tMCDAlarm = Connection.DataTable("ALARMSET, ENVACCELENABLEDANGERHI, ENVACCELENABLEALERTHI, ENVACCELDANGERHI, ENVACCELALERTHI, TEMPERATUREENABLEDANGERHI, TEMPERATUREENABLEALERTHI, TEMPERATUREDANGERHI, TEMPERATUREALERTHI, VELOCITYENABLEDANGERHI, VELOCITYENABLEALERTHI", "MCDAlarm", ((AlarmId > 0) ? ("MCDAlarmId=" + AlarmId.ToString()) : ("ElementId=" + PointId.ToString())));

                            if (tMCDAlarm.Rows.Count > 0)
                            {
                                ReturnValue.Envelope.AlarmMethod = AlarmOverallMethod.Level;
                                ReturnValue.Envelope.EnableDangerHi = (tMCDAlarm.Rows[0]["EnvAccelEnableDangerHi"].ToString() == "1");
                                ReturnValue.Envelope.EnableAlertHi = (tMCDAlarm.Rows[0]["EnvAccelEnableAlertHi"].ToString() == "1");
                                ReturnValue.Envelope.EnableAlertLo = false;
                                ReturnValue.Envelope.EnableDangerLo = false;
                                ReturnValue.Envelope.DangerHi = float.Parse(tMCDAlarm.Rows[0]["EnvAccelDangerHi"].ToString());
                                ReturnValue.Envelope.AlertHi = float.Parse(tMCDAlarm.Rows[0]["EnvAccelAlertHi"].ToString());
                                ReturnValue.Envelope.AlertLo = 0;
                                ReturnValue.Envelope.DangerLo = 0;

                                ReturnValue.Velocity.AlarmMethod = AlarmOverallMethod.Level;
                                ReturnValue.Velocity.EnableDangerHi = (tMCDAlarm.Rows[0]["VelocityEnableDangerHi"].ToString() == "1");
                                ReturnValue.Velocity.EnableAlertHi = (tMCDAlarm.Rows[0]["VelocityEnableAlertHi"].ToString() == "1");
                                ReturnValue.Velocity.EnableAlertLo = false;
                                ReturnValue.Velocity.EnableDangerLo = false;
                                ReturnValue.Velocity.DangerHi = float.Parse(tMCDAlarm.Rows[0]["VelocityDangerHi"].ToString());
                                ReturnValue.Velocity.AlertHi = float.Parse(tMCDAlarm.Rows[0]["VelocityAlertHi"].ToString());
                                ReturnValue.Velocity.AlertLo = 0;
                                ReturnValue.Velocity.DangerLo = 0;

                                ReturnValue.Temperature.AlarmMethod = AlarmOverallMethod.Level;
                                ReturnValue.Temperature.EnableDangerHi = (tMCDAlarm.Rows[0]["TemperatureEnableDangerHi"].ToString() == "1");
                                ReturnValue.Temperature.EnableAlertHi = (tMCDAlarm.Rows[0]["TemperatureEnableAlertHi"].ToString() == "1");
                                ReturnValue.Temperature.EnableAlertLo = false;
                                ReturnValue.Temperature.EnableDangerLo = false;
                                ReturnValue.Temperature.DangerHi = float.Parse(tMCDAlarm.Rows[0]["TemperatureDangerHi"].ToString());
                                ReturnValue.Temperature.AlertHi = float.Parse(tMCDAlarm.Rows[0]["TemperatureAlertHi"].ToString());
                                ReturnValue.Temperature.AlertLo = 0;
                                ReturnValue.Temperature.DangerLo = 0;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + ") error: " + ex.Message);
                    }

                    GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + "): " + ReturnValue.ToString());

                    _AlarmMCD = ReturnValue;
                }
                else
                    _AlarmMCD = null;

                return _AlarmMCD;
            }
        }
        public bool CreateDVA_Percent()
        {
            GenericTools.DebugMsg("CreateDVA_Percent(" + PointId.ToString() + "): Starting");
            bool ReturnValue = false;
            DataTable TreeElemTable;
            uint HierPointId = PointId;
            if (Connection.IsConnected)
            {

                if (PointId > 0)
                {
                    if (Connection.SQLtoInt("count(*)", "TreeElem", "Name='" + TreeElem.Name.Replace("MI ", "DVA ").Replace("OS ", "DVA ") + "' and ParentId=" + TreeElem.ParentId.ToString()) > 0)
                    {
                        ReturnValue = true;
                        GenericTools.DebugMsg("CreateDVA_Percent(" + PointId.ToString() + "): DVA PERCENTUAL point exists");
                    }
                    else
                    {
                        HierPointId = (TreeElem.HierarchyType == HierarchyType.Hierarchy ? PointId : TreeElem.ReferenceId);
                        if (HierPointId > 0)
                        {
                            uint NewTreeElemId = 0;
                            switch (Connection.DBType)
                            {
                                case DBType.Oracle: //Oracle
                                    Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + TreeElem.Name.Replace("MI ", "DVA ").Replace("OS ", "DVA ") + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " Variacao Média', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        );
                                    //NewTreeElemId = Connection.SQLtoUInt(Connection.Owner + "TreeElemId_Seq.CurrVal", "Dual");
                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    Connection.SQLExec("insert into ScalarAlarm (ScalarAlrmId, ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (ScalarAlrmId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80)");
                                    Connection.SQLExec("insert into AlarmAssign (AssignId, AlarmId, ElementId, AlarmIndex, Type) values (AssignId_Seq.NextVal, ScalarAlrmId_Seq.CurrVal, TreeElemId_Seq.CurrVal, 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    Connection.SQLExec("insert into DPExpressionDef (ExpressionId, OwnerId, Name) values (DPExpressionDefId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private expression>')");
                                    Connection.SQLExec("insert into DPExpressionAssign (AssignId, DPId, ExpId) values (DPExpressionAssignId_Seq.NextVal, TreeElemId_Seq.CurrVal, DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionVar (VarKey, VarName, VarType, ExpId) values (DPExpressionVarId_Seq.NextVal, 'Var_" + TreeElem.Name.Replace(" ", "_") + "', 1, DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionVarRef (VarRefId, VarKey, DPId, SourcePtId, ExpId) values (DPExpressionVarRefId_Seq.NextVal, DPExpressionVarId_Seq.CurrVal, TreeElemId_Seq.CurrVal, " + PointId.ToString() + ",  DPExpressionDefId_Seq.CurrVal)");

                                    //Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, DPExpressionDefId_Seq.CurrVal, 1002, 3009)");
                                    //Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, DPExpressionDefId_Seq.CurrVal, 1001, 2006)");
                                    //Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, DPExpressionDefId_Seq.CurrVal, 1005, 1)");
                                    //Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, DPExpressionDefId_Seq.CurrVal, 1003, DPExpressionVarId_Seq.CurrVal)");
                                    //Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, DPExpressionDefId_Seq.CurrVal, 1001, 2007)");





                                    TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    if (TreeElemTable.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                        {
                                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                                + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                                + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                + NewTreeElemId.ToString() + ", " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );

                                            Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                        }
                                    }
                                    break;

                                case DBType.MSSQL:
                                    NewTreeElemId = Connection.SQLtoUInt("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + TreeElem.Name.Replace("MI ", "DVA ").Replace("OS ", "DVA ") + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " Variacao Média', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        + ";select @@Identity;"
                                        );

                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    uint ScalarAlrmId = Connection.SQLtoUInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80);select @@Identity;");
                                    ScalarAlrmId = Connection.SQLtoUInt("SELECT MAX(ScalarAlrmId) from ScalarAlarm");
                                    Connection.SQLExec("insert into AlarmAssign (AlarmId, ElementId, AlarmIndex, Type) values (" + ScalarAlrmId.ToString() + ", " + NewTreeElemId.ToString() + ", 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    uint ExpressionId = Connection.SQLtoUInt("insert into DPExpressionDef (OwnerId, Name) values (" + NewTreeElemId.ToString() + ", '<Private expression>');select @@Identity;");
                                    ExpressionId = Connection.SQLtoUInt("SELECT MAX(ExpressionId) from DPExpressionDef");
                                    Connection.SQLExec("insert into DPExpressionAssign (DPId, ExpId) values (" + NewTreeElemId.ToString() + ", " + ExpressionId.ToString() + ")");

                                    uint VarKey = Connection.SQLtoUInt("insert into DPExpressionVar (VarName, VarType, ExpId) values ('DVA_" + TreeElem.Name.Replace(" ", "_") + "', 1, " + ExpressionId.ToString() + ");select @@Identity;");
                                    VarKey = Connection.SQLtoUInt("SELECT MAX(VarKey) from DPExpressionVar");

                                    uint VarRefId = Connection.SQLtoUInt("insert into DPExpressionVarRef (VarKey, DPId, SourcePtId, ExpId) values (" + VarKey.ToString() + ", " + NewTreeElemId.ToString() + ", " + PointId.ToString() + ", " + ExpressionId.ToString() + ")");
                                    VarRefId = Connection.SQLtoUInt("SELECT MAX(VarRefId) from DPExpressionVarRef");

                                    uint Constante100 = Connection.SQLtoUInt("SELECT CONSTKEY FROM DPCONSTANT WHERE NAME='100' AND EXPID=" + ExpressionId + ")");
                                    if (Constante100 == 0) Connection.SQLtoUInt("INSERT INTO DPCONSTANT (NAME,VALUE,EXPID) VALUES ('100','100'," + ExpressionId + ")");
                                    Constante100 = Connection.SQLtoUInt("SELECT CONSTKEY FROM DPCONSTANT WHERE NAME='100' AND EXPID=" + ExpressionId);

                                    uint Constante1 = Connection.SQLtoUInt("SELECT CONSTKEY FROM DPCONSTANT WHERE NAME='1' AND EXPID=" + ExpressionId + ")");
                                    if (Constante1 == 0) Connection.SQLtoUInt("INSERT INTO DPCONSTANT (NAME,VALUE,EXPID) VALUES ('1','1'," + ExpressionId + ")");
                                    Constante1 = Connection.SQLtoUInt("SELECT CONSTKEY FROM DPCONSTANT WHERE NAME='1' AND EXPID=" + ExpressionId);

                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, " + ExpressionId.ToString() + ",1001,2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, " + ExpressionId.ToString() + ",1001,2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, " + ExpressionId.ToString() + ",1001,2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, " + ExpressionId.ToString() + ",1005,1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, " + ExpressionId.ToString() + ",1003," + VarKey + ")");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (6, " + ExpressionId.ToString() + ",1005,1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (7, " + ExpressionId.ToString() + ",1001,2004)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (8, " + ExpressionId.ToString() + ",1005,1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (9, " + ExpressionId.ToString() + ",1002,3035)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (10, " + ExpressionId.ToString() + ",1001,2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (11, " + ExpressionId.ToString() + ",1003," + VarKey + ")");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (12, " + ExpressionId.ToString() + ",1005,1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (13, " + ExpressionId.ToString() + ",1001,2008)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (14, " + ExpressionId.ToString() + ",1005,1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (15, " + ExpressionId.ToString() + ",1004,28)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (16, " + ExpressionId.ToString() + ",1001,2007)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (17, " + ExpressionId.ToString() + ",1001,2007)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (18, " + ExpressionId.ToString() + ",1005,1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (19, " + ExpressionId.ToString() + ",1001,2002)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (20, " + ExpressionId.ToString() + ",1004," + Constante1 + ")");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (21, " + ExpressionId.ToString() + ",1001,2007)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (22, " + ExpressionId.ToString() + ",1005,1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (23, " + ExpressionId.ToString() + ",1001,2003)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (24, " + ExpressionId.ToString() + ",1005,1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (25, " + ExpressionId.ToString() + ",1004," + Constante100 + ")");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (26, " + ExpressionId.ToString() + ",1001,2007)");




                                    TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    if (TreeElemTable.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                        {
                                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                                + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString().Replace("MI ", "DVA ").Replace("OS ", "DVA ") + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao Média', " //Description
                                                + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                + NewTreeElemId.ToString() + ", " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );

                                            Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                        }
                                    }
                                    break;
                            }
                            ReturnValue = true;
                        }
                    }
                }
            }
            GenericTools.DebugMsg("CreateDVA_Percent(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        public bool CreateDV_HAI(string Rolamento, string Frequencia, string Constante)
        {
            GenericTools.DebugMsg("CreateDV_HAI(" + PointId.ToString() + "): Starting");
            bool ReturnValue = false;
            DataTable TreeElemTable;
            uint HierPointId = PointId;
            if (Connection.IsConnected)
            {

                if (PointId > 0)
                {
                    if (Connection.SQLtoInt("count(*)", "TreeElem", "Name='" + TreeElem.Name.Replace("MI ", "DV HAI ").Replace("OS ", "DV HAI ") + " " + Rolamento + " " + Frequencia + "' and ParentId=" + TreeElem.ParentId.ToString()) > 0)
                    {
                        ReturnValue = true;
                        GenericTools.DebugMsg("CreateDV_HAI(" + PointId.ToString() + "): HAI point exists");
                    }
                    else
                    {
                        HierPointId = (TreeElem.HierarchyType == HierarchyType.Hierarchy ? PointId : TreeElem.ReferenceId);
                        if (HierPointId > 0)
                        {
                            uint NewTreeElemId = 0;
                            switch (Connection.DBType)
                            {

                                case DBType.MSSQL:
                                    NewTreeElemId = Connection.SQLtoUInt("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + TreeElem.Name.Replace("MI ", "DV HAI ").Replace("OS ", "DV HAI ") + " " + Rolamento + " " + Frequencia + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " DERIVADA DE HAI " + Rolamento + " - " + Frequencia + "', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        + ";select @@Identity;"
                                        );

                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, 'HAI')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    uint ScalarAlrmId = Connection.SQLtoUInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 0, 0, 1, 0, 4, 0, 0);select @@Identity;");
                                    ScalarAlrmId = Connection.SQLtoUInt("SELECT MAX(ScalarAlrmId) from ScalarAlarm");
                                    Connection.SQLExec("insert into AlarmAssign (AlarmId, ElementId, AlarmIndex, Type) values (" + ScalarAlrmId.ToString() + ", " + NewTreeElemId.ToString() + ", 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    uint ExpressionId = Connection.SQLtoUInt("insert into DPExpressionDef (OwnerId, Name) values (" + NewTreeElemId.ToString() + ", '<Private expression>');select @@Identity;");
                                    ExpressionId = Connection.SQLtoUInt("SELECT MAX(ExpressionId) from DPExpressionDef");
                                    Connection.SQLExec("insert into DPExpressionAssign (DPId, ExpId) values (" + NewTreeElemId.ToString() + ", " + ExpressionId.ToString() + ")");

                                    uint VarKey = Connection.SQLtoUInt("insert into DPExpressionVar (VarName, VarType, ExpId) values ('" + Frequencia + "', 10, " + ExpressionId.ToString() + ");select @@Identity;");
                                    VarKey = Connection.SQLtoUInt("SELECT MAX(VarKey) from DPExpressionVar");

                                    uint VarRefId = Connection.SQLtoUInt("insert into DPExpressionVarRef (VarKey, DPId, SourcePtId, ExpId) values (" + VarKey.ToString() + ", " + NewTreeElemId.ToString() + ", " + PointId.ToString() + ", " + ExpressionId.ToString() + ")");
                                    VarRefId = Connection.SQLtoUInt("SELECT MAX(VarRefId) from DPExpressionVarRef");

                                    //float Constante_float = float.Parse(Constante);
                                    uint Constante1 = Connection.SQLtoUInt("INSERT INTO DPCONSTANT (NAME,VALUE,EXPID) VALUES ('" + Constante.ToString().Replace(",", ".") + "','" + Constante.ToString().Replace(",", ".") + "'," + ExpressionId + ")");
                                    Constante1 = Connection.SQLtoUInt("SELECT CONSTKEY FROM DPCONSTANT WHERE NAME='" + Constante.ToString().Replace(",", ".") + "' AND EXPID=" + ExpressionId);


                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, " + ExpressionId.ToString() + ", 1002, 3018)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, " + ExpressionId.ToString() + ", 1001, 2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, " + ExpressionId.ToString() + ", 1003, " + VarKey.ToString() + ")");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, " + ExpressionId.ToString() + ", 1001, 2008)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, " + ExpressionId.ToString() + ", 1005, 1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (6, " + ExpressionId.ToString() + ", 1004, '" + Constante1 + "')");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (7, " + ExpressionId.ToString() + ", 1001, 2007)");


                                    TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    if (TreeElemTable.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                        {
                                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                                + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + TreeElem.Name.Replace("MI ", "DV HAI ").Replace("OS ", "DV HAI ") + " " + Rolamento + " " + Frequencia + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " DERIVADA DE HAI " + Rolamento + " - " + Frequencia + "', " //Description
                                                + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                + NewTreeElemId.ToString() + ", " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );

                                            Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                        }
                                    }
                                    break;
                            }
                            ReturnValue = true;
                        }
                    }
                }
            }
            GenericTools.DebugMsg("CreateDV_HAI(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        public bool CreateDVVar()
        {
            GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): Starting");
            bool ReturnValue = false;
            DataTable TreeElemTable;
            uint HierPointId = PointId;
            if (Connection.IsConnected)
            {

                if (PointId > 0)
                {
                    if (Connection.SQLtoInt("count(*)", "TreeElem", "Name='" + TreeElem.Name.Replace("MI ", "DV ").Replace("OS ", "DV ") + "' and ParentId=" + TreeElem.ParentId.ToString()) > 0)
                    {
                        ReturnValue = true;
                        GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): DVVar point exists");
                    }
                    else
                    {
                        HierPointId = (TreeElem.HierarchyType == HierarchyType.Hierarchy ? PointId : TreeElem.ReferenceId);
                        if (HierPointId > 0)
                        {
                            uint NewTreeElemId = 0;
                            switch (Connection.DBType)
                            {
                                #region conexao oracle
                                case DBType.Oracle: //Oracle
                                    Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + TreeElem.Name.Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " Variacao', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        );
                                    //NewTreeElemId = Connection.SQLtoUInt(Connection.Owner + "TreeElemId_Seq.CurrVal", "Dual");
                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    Connection.SQLExec("insert into ScalarAlarm (ScalarAlrmId, ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (ScalarAlrmId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80)");
                                    Connection.SQLExec("insert into AlarmAssign (AssignId, AlarmId, ElementId, AlarmIndex, Type) values (AssignId_Seq.NextVal, ScalarAlrmId_Seq.CurrVal, TreeElemId_Seq.CurrVal, 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    Connection.SQLExec("insert into DPExpressionDef (ExpressionId, OwnerId, Name) values (DPExpressionDefId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private expression>')");
                                    Connection.SQLExec("insert into DPExpressionAssign (AssignId, DPId, ExpId) values (DPExpressionAssignId_Seq.NextVal, TreeElemId_Seq.CurrVal, DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionVar (VarKey, VarName, VarType, ExpId) values (DPExpressionVarId_Seq.NextVal, 'Var_" + TreeElem.Name.Replace(" ", "_") + "', 1, DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionVarRef (VarRefId, VarKey, DPId, SourcePtId, ExpId) values (DPExpressionVarRefId_Seq.NextVal, DPExpressionVarId_Seq.CurrVal, TreeElemId_Seq.CurrVal, " + PointId.ToString() + ",  DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, DPExpressionDefId_Seq.CurrVal, 1002, 3009)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, DPExpressionDefId_Seq.CurrVal, 1001, 2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, DPExpressionDefId_Seq.CurrVal, 1005, 1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, DPExpressionDefId_Seq.CurrVal, 1003, DPExpressionVarId_Seq.CurrVal)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, DPExpressionDefId_Seq.CurrVal, 1001, 2007)");
                                    TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    if (TreeElemTable.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                        {
                                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                                + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                                + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                + NewTreeElemId.ToString() + ", " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );

                                            Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                        }
                                    }
                                    break;
                                #endregion

                                case DBType.MSSQL:
                                    NewTreeElemId = Connection.SQLtoUInt("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + TreeElem.Name.Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " Variacao', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        + ";select @@Identity;"
                                        );

                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    uint ScalarAlrmId = Connection.SQLtoUInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80);select @@Identity;");
                                    ScalarAlrmId = Connection.SQLtoUInt("SELECT MAX(ScalarAlrmId) from ScalarAlarm");
                                    Connection.SQLExec("insert into AlarmAssign (AlarmId, ElementId, AlarmIndex, Type) values (" + ScalarAlrmId.ToString() + ", " + NewTreeElemId.ToString() + ", 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    uint ExpressionId = Connection.SQLtoUInt("insert into DPExpressionDef (OwnerId, Name) values (" + NewTreeElemId.ToString() + ", '<Private expression>');select @@Identity;");
                                    ExpressionId = Connection.SQLtoUInt("SELECT MAX(ExpressionId) from DPExpressionDef");
                                    Connection.SQLExec("insert into DPExpressionAssign (DPId, ExpId) values (" + NewTreeElemId.ToString() + ", " + ExpressionId.ToString() + ")");

                                    uint VarKey = Connection.SQLtoUInt("insert into DPExpressionVar (VarName, VarType, ExpId) values ('Var_" + TreeElem.Name.Replace(" ", "_") + "', 1, " + ExpressionId.ToString() + ");select @@Identity;");
                                    VarKey = Connection.SQLtoUInt("SELECT MAX(VarKey) from DPExpressionVar");

                                    uint VarRefId = Connection.SQLtoUInt("insert into DPExpressionVarRef (VarKey, DPId, SourcePtId, ExpId) values (" + VarKey.ToString() + ", " + NewTreeElemId.ToString() + ", " + PointId.ToString() + ", " + ExpressionId.ToString() + ")");
                                    VarRefId = Connection.SQLtoUInt("SELECT MAX(VarRefId) from DPExpressionVarRef");


                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, " + ExpressionId.ToString() + ", 1002, 3009)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, " + ExpressionId.ToString() + ", 1001, 2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, " + ExpressionId.ToString() + ", 1005, 1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, " + ExpressionId.ToString() + ", 1003, " + VarKey.ToString() + ")");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, " + ExpressionId.ToString() + ", 1001, 2007)");
                                    //TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    //if (TreeElemTable.Rows.Count > 0)
                                    //{
                                    //    for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                    //    {
                                    //        Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                    //            + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                    //            + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                    //            + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                    //            + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                    //            + "'" + TreeElemTable.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                    //            + "4, " //ContainerType
                                    //            + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                    //            + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                    //            + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                    //            + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                    //            + "1, " //AlarmFlags
                                    //            + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                    //            + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                    //            + NewTreeElemId.ToString() + ", " //ReferenceId
                                    //            + "0, " //Good
                                    //            + "0, " //Alert
                                    //            + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                    //            );

                                    //        Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                    //    }
                                    //}
                                    break;
                            }
                            ReturnValue = true;
                        }
                    }
                }
            }
            GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public bool CreateDVVar_WithoutMI()
        {
            GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): Starting");
            bool ReturnValue = false;
            DataTable TreeElemTable;
            uint HierPointId = PointId;
            if (Connection.IsConnected)
            {

                if (PointId > 0)
                {
                    if (Connection.SQLtoInt("count(*)", "TreeElem", "Name='DV " + TreeElem.Name + "' and ParentId=" + TreeElem.ParentId.ToString()) > 0)
                    {
                        ReturnValue = true;
                        GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): DVVar point exists");
                    }
                    else
                    {
                        HierPointId = (TreeElem.HierarchyType == HierarchyType.Hierarchy ? PointId : TreeElem.ReferenceId);
                        if (HierPointId > 0)
                        {
                            uint NewTreeElemId = 0;
                            switch (Connection.DBType)
                            {
                                #region conexao oracle
                                case DBType.Oracle: //Oracle
                                    Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'DV " + TreeElem.Name + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " Variacao', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        );
                                    //NewTreeElemId = Connection.SQLtoUInt(Connection.Owner + "TreeElemId_Seq.CurrVal", "Dual");
                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    Connection.SQLExec("insert into ScalarAlarm (ScalarAlrmId, ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (ScalarAlrmId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80)");
                                    Connection.SQLExec("insert into AlarmAssign (AssignId, AlarmId, ElementId, AlarmIndex, Type) values (AssignId_Seq.NextVal, ScalarAlrmId_Seq.CurrVal, TreeElemId_Seq.CurrVal, 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    Connection.SQLExec("insert into DPExpressionDef (ExpressionId, OwnerId, Name) values (DPExpressionDefId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private expression>')");
                                    Connection.SQLExec("insert into DPExpressionAssign (AssignId, DPId, ExpId) values (DPExpressionAssignId_Seq.NextVal, TreeElemId_Seq.CurrVal, DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionVar (VarKey, VarName, VarType, ExpId) values (DPExpressionVarId_Seq.NextVal, 'Var_" + TreeElem.Name.Replace(" ", "_") + "', 1, DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionVarRef (VarRefId, VarKey, DPId, SourcePtId, ExpId) values (DPExpressionVarRefId_Seq.NextVal, DPExpressionVarId_Seq.CurrVal, TreeElemId_Seq.CurrVal, " + PointId.ToString() + ",  DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, DPExpressionDefId_Seq.CurrVal, 1002, 3009)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, DPExpressionDefId_Seq.CurrVal, 1001, 2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, DPExpressionDefId_Seq.CurrVal, 1005, 1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, DPExpressionDefId_Seq.CurrVal, 1003, DPExpressionVarId_Seq.CurrVal)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, DPExpressionDefId_Seq.CurrVal, 1001, 2007)");
                                    TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    if (TreeElemTable.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                        {
                                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                                + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                                + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                + NewTreeElemId.ToString() + ", " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );

                                            Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                        }
                                    }
                                    break;
                                #endregion

                                case DBType.MSSQL:
                                    NewTreeElemId = Connection.SQLtoUInt("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'DV " + TreeElem.Name + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " Variacao', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        + ";select @@Identity;"
                                        );

                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    uint ScalarAlrmId = 0;
                                    if (PointType == PointType.AccEnvelope)
                                        ScalarAlrmId = Connection.SQLtoUInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 1, 0, 1, 0, 500, 0, 120);SELECT SCOPE_IDENTITY()");
                                    else if (PointType == PointType.AccToVel)
                                        ScalarAlrmId = Connection.SQLtoUInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 1, 0, 1, 0, 500, 0, 200);SELECT SCOPE_IDENTITY()");
                                    else
                                        ScalarAlrmId = Connection.SQLtoUInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80);SELECT SCOPE_IDENTITY()");
                                    Connection.SQLExec("insert into AlarmAssign (AlarmId, ElementId, AlarmIndex, Type) values (" + ScalarAlrmId.ToString() + ", " + NewTreeElemId.ToString() + ", 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    uint ExpressionId = Connection.SQLtoUInt("insert into DPExpressionDef (OwnerId, Name) values (" + NewTreeElemId.ToString() + ", '<Private expression>');select @@Identity;");
                                    ExpressionId = Connection.SQLtoUInt("SELECT MAX(ExpressionId) from DPExpressionDef");
                                    Connection.SQLExec("insert into DPExpressionAssign (DPId, ExpId) values (" + NewTreeElemId.ToString() + ", " + ExpressionId.ToString() + ")");

                                    uint VarKey = Connection.SQLtoUInt("insert into DPExpressionVar (VarName, VarType, ExpId) values ('Var_" + TreeElem.Name.Replace(" ", "_") + "', 1, " + ExpressionId.ToString() + ");select @@Identity;");
                                    VarKey = Connection.SQLtoUInt("SELECT MAX(VarKey) from DPExpressionVar");

                                    uint VarRefId = Connection.SQLtoUInt("insert into DPExpressionVarRef (VarKey, DPId, SourcePtId, ExpId) values (" + VarKey.ToString() + ", " + NewTreeElemId.ToString() + ", " + PointId.ToString() + ", " + ExpressionId.ToString() + ")");
                                    VarRefId = Connection.SQLtoUInt("SELECT MAX(VarRefId) from DPExpressionVarRef");


                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, " + ExpressionId.ToString() + ", 1002, 3009)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, " + ExpressionId.ToString() + ", 1001, 2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, " + ExpressionId.ToString() + ", 1005, 1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, " + ExpressionId.ToString() + ", 1003, " + VarKey.ToString() + ")");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, " + ExpressionId.ToString() + ", 1001, 2007)");
                                    //TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    //if (TreeElemTable.Rows.Count > 0)
                                    //{
                                    //    for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                    //    {
                                    //        Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                    //            + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                    //            + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                    //            + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                    //            + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                    //            + "'" + TreeElemTable.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                    //            + "4, " //ContainerType
                                    //            + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                    //            + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                    //            + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                    //            + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                    //            + "1, " //AlarmFlags
                                    //            + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                    //            + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                    //            + NewTreeElemId.ToString() + ", " //ReferenceId
                                    //            + "0, " //Good
                                    //            + "0, " //Alert
                                    //            + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                    //            );

                                    //        Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                    //    }
                                    //}
                                    break;
                            }
                            ReturnValue = true;
                        }
                    }
                }
            }
            GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public bool CreateDVVar_SEMPADRAO_SSCP()
        {
            GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): Starting");
            bool ReturnValue = false;
            DataTable TreeElemTable;
            uint HierPointId = PointId;
            if (Connection.IsConnected)
            {

                if (PointId > 0)
                {
                    if (Connection.SQLtoInt("count(*)", "TreeElem", "DESCRIPTION='" + TreeElem.Name + " Variacao' and ParentId=" + TreeElem.ParentId.ToString()) > 0)
                    {
                        ReturnValue = true;
                        GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): DVVar point exists");
                    }
                    else
                    {
                        HierPointId = (TreeElem.HierarchyType == HierarchyType.Hierarchy ? PointId : TreeElem.ReferenceId);
                        if (HierPointId > 0)
                        {
                            uint NewTreeElemId = 0;
                            switch (Connection.DBType)
                            {
                                #region Oracle
                                case DBType.Oracle: //Oracle
                                    Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + TreeElem.Name.Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " Variacao', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        );
                                    //NewTreeElemId = Connection.SQLtoUInt(Connection.Owner + "TreeElemId_Seq.CurrVal", "Dual");
                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    Connection.SQLExec("insert into ScalarAlarm (ScalarAlrmId, ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (ScalarAlrmId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80)");
                                    Connection.SQLExec("insert into AlarmAssign (AssignId, AlarmId, ElementId, AlarmIndex, Type) values (AssignId_Seq.NextVal, ScalarAlrmId_Seq.CurrVal, TreeElemId_Seq.CurrVal, 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    Connection.SQLExec("insert into DPExpressionDef (ExpressionId, OwnerId, Name) values (DPExpressionDefId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private expression>')");
                                    Connection.SQLExec("insert into DPExpressionAssign (AssignId, DPId, ExpId) values (DPExpressionAssignId_Seq.NextVal, TreeElemId_Seq.CurrVal, DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionVar (VarKey, VarName, VarType, ExpId) values (DPExpressionVarId_Seq.NextVal, 'Var_" + TreeElem.Name.Replace(" ", "_") + "', 1, DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionVarRef (VarRefId, VarKey, DPId, SourcePtId, ExpId) values (DPExpressionVarRefId_Seq.NextVal, DPExpressionVarId_Seq.CurrVal, TreeElemId_Seq.CurrVal, " + PointId.ToString() + ",  DPExpressionDefId_Seq.CurrVal)");

                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, DPExpressionDefId_Seq.CurrVal, 1002, 3009)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, DPExpressionDefId_Seq.CurrVal, 1001, 2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, DPExpressionDefId_Seq.CurrVal, 1005, 1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, DPExpressionDefId_Seq.CurrVal, 1003, DPExpressionVarId_Seq.CurrVal)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, DPExpressionDefId_Seq.CurrVal, 1001, 2007)");
                                    TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    if (TreeElemTable.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                        {
                                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                                + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                                + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                + NewTreeElemId.ToString() + ", " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );

                                            Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                        }
                                    }
                                    break;
                                #endregion

                                case DBType.MSSQL:
                                    int lengh = TreeElem.Name.Length;

                                    string Nome = (TreeElem.Name.Substring(0, (lengh > 1 ? lengh : 1)) == "MI" ? TreeElem.Name.Replace("MI ", "DV ") : "DV " + TreeElem.Name.Substring(0, (lengh > 17 ? 17 : lengh)));

                                    NewTreeElemId = Connection.SQLtoUInt("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + Nome + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + TreeElem.Name + " Variacao', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        + ";select @@Identity;"
                                        );

                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_DerivedPoint") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASPT_DerivedPointCalculated") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, '%')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, '" + Connection.SQLtoString("DefaultName", "Registration", "Signature='SKFCM_ASPT_DerivedPointCalculated'") + "')");

                                    uint ScalarAlrmId = Connection.SQLtoUInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 1, 0, 1, 0, 100, 0, 80);select @@Identity;");
                                    ScalarAlrmId = Connection.SQLtoUInt("SELECT MAX(ScalarAlrmId) from ScalarAlarm");
                                    Connection.SQLExec("insert into AlarmAssign (AlarmId, ElementId, AlarmIndex, Type) values (" + ScalarAlrmId.ToString() + ", " + NewTreeElemId.ToString() + ", 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");

                                    uint ExpressionId = Connection.SQLtoUInt("insert into DPExpressionDef (OwnerId, Name) values (" + NewTreeElemId.ToString() + ", '<Private expression>');select @@Identity;");
                                    ExpressionId = Connection.SQLtoUInt("SELECT MAX(ExpressionId) from DPExpressionDef");
                                    Connection.SQLExec("insert into DPExpressionAssign (DPId, ExpId) values (" + NewTreeElemId.ToString() + ", " + ExpressionId.ToString() + ")");

                                    uint VarKey = Connection.SQLtoUInt("insert into DPExpressionVar (VarName, VarType, ExpId) values ('Var_" + TreeElem.Name.Replace(" ", "_") + "', 1, " + ExpressionId.ToString() + ");select @@Identity;");
                                    VarKey = Connection.SQLtoUInt("SELECT MAX(VarKey) from DPExpressionVar");

                                    uint VarRefId = Connection.SQLtoUInt("insert into DPExpressionVarRef (VarKey, DPId, SourcePtId, ExpId) values (" + VarKey.ToString() + ", " + NewTreeElemId.ToString() + ", " + PointId.ToString() + ", " + ExpressionId.ToString() + ")");
                                    VarRefId = Connection.SQLtoUInt("SELECT MAX(VarRefId) from DPExpressionVarRef");


                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (1, " + ExpressionId.ToString() + ", 1002, 3009)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (2, " + ExpressionId.ToString() + ", 1001, 2006)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (3, " + ExpressionId.ToString() + ", 1005, 1)");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (4, " + ExpressionId.ToString() + ", 1003, " + VarKey.ToString() + ")");
                                    Connection.SQLExec("insert into DPExpressionFormula (OrderIndex, ExpId, ItemType, ItemKey) values (5, " + ExpressionId.ToString() + ", 1001, 2007)");
                                    //TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    //if (TreeElemTable.Rows.Count > 0)
                                    //{
                                    //    for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                    //    {
                                    //        Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                    //            + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                    //            + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                    //            + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                    //            + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                    //            + "'" + TreeElemTable.Rows[i]["Name"].ToString().Replace("MI ", "DV ").Replace("OS ", "DV ") + "', " //Name
                                    //            + "4, " //ContainerType
                                    //            + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                    //            + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                    //            + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                    //            + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                    //            + "1, " //AlarmFlags
                                    //            + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                    //            + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                    //            + NewTreeElemId.ToString() + ", " //ReferenceId
                                    //            + "0, " //Good
                                    //            + "0, " //Alert
                                    //            + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                    //            );

                                    //        Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                    //    }
                                    //}
                                    break;
                            }
                            ReturnValue = true;
                        }
                    }
                }
            }
            GenericTools.DebugMsg("CreateDVVar(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public bool CopyPoint(string New_Name, PointType Type)
        {


            GenericTools.DebugMsg("CopyPoint(" + PointId.ToString() + "): Starting");
            bool ReturnValue = false;
            DataTable TreeElemTable;
            uint HierPointId = PointId;
            if (Connection.IsConnected)
            {

                if (PointId > 0)
                {
                    if (Connection.SQLtoInt("count(*)", "TreeElem", "Name='" + New_Name + "' and ParentId=" + TreeElem.ParentId.ToString()) > 0)
                    {
                        ReturnValue = true;
                        GenericTools.DebugMsg("CopyPoint(" + PointId.ToString() + "): This point exists");
                    }
                    else
                    {
                        string Registration_Type_Id = "";
                        switch (Type)
                        {
                            case Analyst.PointType.AccToVel:
                                Registration_Type_Id = "SKFCM_ASPT_AccToVel";
                                break;

                        }


                        HierPointId = (TreeElem.HierarchyType == HierarchyType.Hierarchy ? PointId : TreeElem.ReferenceId);
                        if (HierPointId > 0)
                        {
                            uint NewTreeElemId = 0;
                            switch (Connection.DBType)
                            {
                                case DBType.Oracle: //Oracle
                                    Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        + (TreeElem.SlotNumber + 1) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + New_Name + "', " //Name
                                        + "4, " //ContainerType
                                        + "'" + New_Name.Substring(3, 3) + "', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        );
                                    //NewTreeElemId = Connection.SQLtoUInt(Connection.Owner + "TreeElemId_Seq.CurrVal", "Dual");
                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    //Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    //Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_MicrologDAD") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, Registration_Type_Id) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, 'mm/s')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + Connection.Owner + "TreeElemId_Seq.CurrVal, " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, 'Acelerômetro')");

                                    Connection.SQLExec("insert into ScalarAlarm (ScalarAlrmId, ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (ScalarAlrmId_Seq.NextVal, TreeElemId_Seq.CurrVal, '<Private alarm>', 1, 0, 1, 0, 1, 0, " + this.DangerHi + ", 0, " + this.AlertHi + ")");
                                    Connection.SQLExec("insert into AlarmAssign (AssignId, AlarmId, ElementId, AlarmIndex, Type) values (AssignId_Seq.NextVal, ScalarAlrmId_Seq.CurrVal, TreeElemId_Seq.CurrVal, 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");


                                    TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    if (TreeElemTable.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                        {
                                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (TreeElemId, HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (TreeElemId_Seq.NextVal, " //TreeElemId
                                                + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + New_Name + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                                + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                + NewTreeElemId.ToString() + ", " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );

                                            Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                        }
                                    }
                                    break;

                                case DBType.MSSQL:
                                    NewTreeElemId = Connection.SQLtoUInt("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                        + TreeElem.HierarchyId.ToString() + ", " //HierarchyId
                                        + TreeElem.BranchLevel.ToString() + ", " //BranchLevel
                                        //+ Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElem.ParentId.ToString()) + ", " //SlotNumber
                                        + (TreeElem.SlotNumber + 1) + ", " //SlotNumber
                                        + TreeElem.TblSetId.ToString() + ", " //TblSetId
                                        + "'" + New_Name + "', " //Name
                                        + "4, " //ContainerType
                                        + "'@" + New_Name.Substring(3, 2) + "', " //Description
                                        + (TreeElem.ElementEnable ? "1, " : "0, ") //ElementEnable
                                        + (TreeElem.ParentEnable ? "1, " : "0, ") //ParentEnable
                                        + "1, " //HierarchyType
                                        + "1, " //AlarmFlags
                                        + TreeElem.ParentId.ToString() + ", " //ParentId
                                        + TreeElem.ParentId.ToString() + ", " //ParentRefId
                                        + "0, " //ReferenceId
                                        + "0, " //Good
                                        + "0, " //Alert
                                        + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                        + ";select @@Identity;"
                                        );


                                    NewTreeElemId = Connection.SQLtoUInt("SELECT MAX(TREEELEMID) FROM TREEELEM");
                                    //Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTime") + ", 1, '900')");
                                    //Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_DPEvalTimeUnits") + ", 1, '21501')"); 
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Dad_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASDD_MicrologDAD") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Type_Id") + ", 1, '" + Registration.RegistrationId(Connection, Registration_Type_Id) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale") + ", 2, '5')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit") + ", 3, 'mm/s')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASUT_NoUnitsEnglish") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Location") + ", 3, '')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Orientation") + ", 3, 'None')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Point_Units") + ", 1, '0')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + ", 1, '" + Connection.SQLtoString("ValueString", "Point", "FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit") + " and ElementId=" + PointId.ToString()) + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Application_Id") + ", 1, '" + Registration.RegistrationId(Connection, "SKFCM_ASAS_General") + "')");
                                    Connection.SQLExec("insert into " + Connection.Owner + "Point (ElementId, FieldId, DataType, ValueString) values (" + NewTreeElemId.ToString() + ", " + Registration.RegistrationId(Connection, "SKFCM_ASPF_Sensor") + ", 1, 'Acelerômetro')");

                                    uint ScalarAlrmId = Connection.SQLtoUInt("insert into ScalarAlarm (ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi) values (" + NewTreeElemId.ToString() + ", '<Private alarm>', 1, 0, 1, 0, 1, 0, " + this.DangerHi + ", 0, " + this.AlertHi + ");select @@Identity;");
                                    ScalarAlrmId = Connection.SQLtoUInt("SELECT MAX(ScalarAlrmId) from ScalarAlarm");
                                    Connection.SQLExec("insert into AlarmAssign (AlarmId, ElementId, AlarmIndex, Type) values (" + ScalarAlrmId.ToString() + ", " + NewTreeElemId.ToString() + ", 1, " + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall") + ")");


                                    TreeElemTable = Connection.DataTable("*", "TreeElem", "ReferenceId=" + PointId.ToString() + " AND Parentid!=2147000000");
                                    if (TreeElemTable.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < TreeElemTable.Rows.Count; i++)
                                        {
                                            Connection.SQLExec("insert into " + Connection.Owner + "TreeElem (HierarchyId, BranchLevel, SlotNumber, TblSetId, Name, ContainerType, Description, ElementEnable, ParentEnable, HierarchyType, AlarmFlags, ParentId, ParentRefId, ReferenceId, Good, Alert, Danger" + (Connection.Version.StartsWith("4.") ? string.Empty : ", Overdue, ChannelEnable") + ") values (" //TreeElemId
                                                + TreeElemTable.Rows[i]["HierarchyId"].ToString() + ", " //HierarchyId
                                                + TreeElemTable.Rows[i]["BranchLevel"].ToString() + ", " //BranchLevel
                                                + Connection.SQLtoString("1+MAX(SlotNumber)", "TreeElem", "ParentId = " + TreeElemTable.Rows[0]["ParentId"].ToString()) + ", " //SlotNumber
                                                + TreeElemTable.Rows[i]["TblSetId"].ToString() + ", " //TblSetId
                                                + "'" + New_Name + "', " //Name
                                                + "4, " //ContainerType
                                                + "'" + TreeElemTable.Rows[i]["Name"].ToString() + " Variacao', " //Description
                                                + TreeElemTable.Rows[i]["ElementEnable"].ToString() + ", " //ElementEnable
                                                + TreeElemTable.Rows[i]["ParentEnable"].ToString() + ", " //ParentEnable
                                                + TreeElemTable.Rows[i]["HierarchyType"].ToString() + ", " //HierarchyType
                                                + "1, " //AlarmFlags
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentId
                                                + TreeElemTable.Rows[i]["ParentId"].ToString() + ", " //ParentRefId
                                                + NewTreeElemId.ToString() + ", " //ReferenceId
                                                + "0, " //Good
                                                + "0, " //Alert
                                                + (Connection.Version.StartsWith("4.") ? "0)" : "0, 0, 1)") //Danger, Overdue, ChannelEnable
                                                );

                                            Log.log("Parentid=" + TreeElemTable.Rows[i]["ParentId"].ToString() + " HierarchyType= " + TreeElemTable.Rows[i]["HierarchyType"].ToString() + " ReferenceId=" + NewTreeElemId.ToString());
                                        }
                                    }
                                    break;
                            }
                            ReturnValue = true;
                        }
                    }
                }
            }
            GenericTools.DebugMsg("CopyPoint(" + PointId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        private string _FullScaleUnit = string.Empty;
        public string FullScaleUnit
        {
            get
            {
                _FullScaleUnit = (string.IsNullOrEmpty(_FullScaleUnit) && IsLoaded ? ValueString("SKFCM_ASPF_Full_Scale_Unit") : string.Empty);
                return _FullScaleUnit;
            }
        }

        public void FixConditionalPoints()
        {
            if (_TreeElem.HierarchyType == HierarchyType.Hierarchy)
            {
                StringBuilder Query = new StringBuilder();
                Query.Append(" select");
                Query.Append("   Po.ElementId as ElementId");
                Query.Append("   ,TE3.TreeElemId as TreeElemId");
                Query.Append(" from");
                Query.Append("   " + _TreeElem.Connection.Owner + "Point Po");
                Query.Append("   ," + _TreeElem.Connection.Owner + "TreeElem TE1");
                Query.Append("   ," + _TreeElem.Connection.Owner + "TreeElem TE2");
                Query.Append("   ," + _TreeElem.Connection.Owner + "TreeElem TE3");
                Query.Append(" where");
                Query.Append("   and Po.FieldId=" + Registration.RegistrationId(_TreeElem.Connection, "SKFCM_ASPF_Conditional_Point").ToString());
                Query.Append("   and Po.ValueString!='0'");
                Query.Append("   and TE1.TreeElemId=Po.ElementId");
                Query.Append("   and TE1.TreeElemId=" + _TreeElem.TreeElemId.ToString());
                Query.Append("   and TE1.ParentId!=2147000000");
                Query.Append("   and TE2.ParentId!=2147000000");
                Query.Append("   and TE2.ParentId!=TE1.ParentId");
                Query.Append("   and TE2.TreeElemId=to_number(Po.ValueString)");
                Query.Append("   and TE3.ParentId=TE1.ParentId");
                Query.Append("   and TE3.Name=TE2.Name");

                DataTable ConditionalPoints = _TreeElem.Connection.DataTable(Query.ToString());

                if (ConditionalPoints.Rows.Count > 0)
                {
                    foreach (DataRow Row in ConditionalPoints.Rows)
                        _TreeElem.Connection.SQLUpdate("Point", "ValueString", Row["TreeElemId"].ToString(), "ElementId=" + Row["ElementId"].ToString() + " and FieldId=" + Registration.RegistrationId(_TreeElem.Connection, "SKFCM_ASPF_Conditional_Point").ToString());
                    _ConditionalPoint = null;
                }
            }
        }

        private string _StdName = string.Empty;
        public string StdName
        {
            get
            {
                if (string.IsNullOrEmpty(_StdName))
                {
                    string Aux = string.Empty;

                    //DAD           "MI "
                    Aux = DadPrefix;
                    if (!string.IsNullOrEmpty(Aux))
                    {
                        string NameCalc = Aux + " ";

                        //Position      "MI 01"
                        Aux = _PosFromName();
                        if (!string.IsNullOrEmpty(Aux))
                        {
                            if (GenericTools.StrToInt(Aux) < 10) Aux = "0" + (GenericTools.StrToInt(Aux)).ToString();
                            NameCalc += Aux;

                            //Orientation   "MI 01H"
                            Aux = OrientationPrefixFromName;
                            if (Aux != string.Empty)
                            {
                                NameCalc += Aux;

                                //Tech          "MI 01HA"
                                Aux = Tech;
                                if (Aux != string.Empty)
                                {
                                    NameCalc += Aux;

                                    _StdName = NameCalc;
                                }
                            }
                        }
                    }
                }
                return _StdName;
            }
        }

        private DadGroup DadGroup
        {
            get
            {
                switch (DadType)
                {
                    case Dad.DMx:
                    case Dad.Gal:
                    case Dad.Gen:
                    case Dad.IMx:
                    case Dad.IMxM:
                    case Dad.IMxP:
                    case Dad.IMxS:
                    case Dad.IMxT:
                    case Dad.LMU:
                    case Dad.MasCon:
                    case Dad.Mascon16:
                    case Dad.MIM:
                    case Dad.TMU:
                    case Dad.WMx:
                    case Dad.WMx_Sub_WVT:
                        return Analyst.DadGroup.Multilog;

                    case Dad.Microlog: return Analyst.DadGroup.MicrologAnalyzer;
                    case Dad.Marlin: return Analyst.DadGroup.MicrologInspector;
                    case Dad.Manual: return Analyst.DadGroup.ManualEntry;
                    case Dad.DerivedPoint: return Analyst.DadGroup.DerivedPoint;
                    case Dad.OilAnalysis: return Analyst.DadGroup.TrendOil;
                    default: return Analyst.DadGroup.None;
                }
            }
        }

        public string DadPrefix
        {
            get
            {
                switch (DadGroup)
                {
                    case Analyst.DadGroup.Multilog: return "OS";
                    case Analyst.DadGroup.MicrologAnalyzer: return "MI";
                    case Analyst.DadGroup.MicrologInspector: return "MA";
                    case Analyst.DadGroup.ManualEntry: return "ME";
                    case Analyst.DadGroup.TrendOil: return "TO";
                    case Analyst.DadGroup.DerivedPoint: return "DV";
                    case Analyst.DadGroup.EletronicEntry: return "EE";
                    default: return string.Empty;
                }
            }
        }

        private string _PosFromName()
        {
            string sDBStandard_TreeElem_Name_Ref = NameRef();
            string sDBStandard_Pos = string.Empty;
            string sDBStandard_Pos_Ref = string.Empty;
            string sDBStandard_Pos_Ref_TMP = string.Empty;
            int nDBStandard_Pos_Ref = 0;
            int nDBStandard_Pos_Ref1 = 0;
            int nDBStandard_Pos_Ref2 = 0;

            if (sDBStandard_TreeElem_Name_Ref.IndexOf(" ") < 0)
                if (!sDBStandard_TreeElem_Name_Ref.StartsWith(DadPrefix))
                    sDBStandard_TreeElem_Name_Ref = DadPrefix + " " + sDBStandard_TreeElem_Name_Ref;

            if (!string.IsNullOrEmpty(OrientationPrefixFromName))
            {
                nDBStandard_Pos_Ref = sDBStandard_TreeElem_Name_Ref.IndexOf(Orientation + Tech.Substring(0, 1));
                if (nDBStandard_Pos_Ref < 0) nDBStandard_Pos_Ref = sDBStandard_TreeElem_Name_Ref.IndexOf(Orientation + " " + Tech.Substring(0, 1));
                if (nDBStandard_Pos_Ref < 0) nDBStandard_Pos_Ref = sDBStandard_TreeElem_Name_Ref.IndexOf((Tech == "T" ? string.Empty : Orientation + " ") + Tech.Substring(0, 1));
                if (nDBStandard_Pos_Ref < 0)
                {
                    return string.Empty;
                }
                else
                {
                    sDBStandard_Pos_Ref_TMP = (sDBStandard_TreeElem_Name_Ref.Substring(0, nDBStandard_Pos_Ref + 1)).Trim();
                    nDBStandard_Pos_Ref1 = sDBStandard_Pos_Ref_TMP.LastIndexOf(" ");
                    if (nDBStandard_Pos_Ref1 > 0) sDBStandard_Pos_Ref_TMP = (sDBStandard_Pos_Ref_TMP.Substring(nDBStandard_Pos_Ref1)).Trim();
                    nDBStandard_Pos_Ref2 = Math.Max(sDBStandard_Pos_Ref_TMP.IndexOf((Tech == "T" ? string.Empty : Orientation + " ") + Tech.Substring(0, 1)), 1);
                    sDBStandard_Pos_Ref = (sDBStandard_Pos_Ref_TMP.Substring(0, Math.Min(nDBStandard_Pos_Ref2, sDBStandard_Pos_Ref_TMP.Length))).Trim();
                    if (GenericTools.StrToInt(sDBStandard_Pos_Ref) > 0)
                    {
                        if (GenericTools.StrToInt(sDBStandard_Pos_Ref) < 10) sDBStandard_Pos = "0";
                        sDBStandard_Pos += GenericTools.StrToInt(sDBStandard_Pos_Ref).ToString();
                        return sDBStandard_Pos;
                    }
                    else
                    {
                        sDBStandard_Pos_Ref_TMP = (sDBStandard_TreeElem_Name_Ref.Substring(0, nDBStandard_Pos_Ref + 1)).Trim();
                        nDBStandard_Pos_Ref2 = Math.Max(sDBStandard_Pos_Ref_TMP.IndexOf((Tech == "T" ? string.Empty : Orientation + " ") + Tech.Substring(0, 1)), 1);
                        sDBStandard_Pos_Ref = (sDBStandard_Pos_Ref_TMP.Substring(0, Math.Min(nDBStandard_Pos_Ref2 + 1, sDBStandard_Pos_Ref_TMP.Length))).Trim();
                        if (GenericTools.StrToInt(sDBStandard_Pos_Ref) > 0)
                        {
                            if (GenericTools.StrToInt(sDBStandard_Pos_Ref) < 10) sDBStandard_Pos = "0";
                            sDBStandard_Pos += GenericTools.StrToInt(sDBStandard_Pos_Ref).ToString();
                            return sDBStandard_Pos;
                        }
                        else
                        {
                            nDBStandard_Pos_Ref1 = sDBStandard_Pos_Ref_TMP.IndexOf(" ");
                            if (nDBStandard_Pos_Ref1 > 0) sDBStandard_Pos_Ref_TMP = (sDBStandard_Pos_Ref_TMP.Substring(nDBStandard_Pos_Ref1)).Trim();
                            nDBStandard_Pos_Ref2 = Math.Max(sDBStandard_Pos_Ref_TMP.IndexOf((Tech == "T" ? string.Empty : Orientation + " ") + Tech.Substring(0, 1)), 1);
                            sDBStandard_Pos_Ref = (sDBStandard_Pos_Ref_TMP.Substring(0, Math.Min(nDBStandard_Pos_Ref2 + 1, sDBStandard_Pos_Ref_TMP.Length))).Trim();
                            if (GenericTools.StrToInt(sDBStandard_Pos_Ref) > 0)
                            {
                                if (GenericTools.StrToInt(sDBStandard_Pos_Ref) < 10) sDBStandard_Pos = "0";
                                sDBStandard_Pos += GenericTools.StrToInt(sDBStandard_Pos_Ref).ToString();
                                return sDBStandard_Pos;
                            }
                        }
                    }
                }
            }
            return string.Empty;
        }

        private bool _OrientationLoaded = false;
        private Orientation _Orientation;
        public Orientation Orientation
        {
            get
            {
                if (!_OrientationLoaded)
                {
                    switch (ValueString("SKFCM_ASPF_Orientation").ToUpper())
                    {
                        case "NONE":
                        case "NENHUM":
                            _Orientation = Analyst.Orientation.None;
                            break;

                        case "HORIZONTAL":
                            _Orientation = Analyst.Orientation.Horizontal;
                            break;

                        case "VERTICAL":
                            _Orientation = Analyst.Orientation.Vertical;
                            break;

                        case "AXIAL":
                            _Orientation = Analyst.Orientation.Axial;
                            break;

                        case "RADIAL":
                            _Orientation = Analyst.Orientation.Radial;
                            break;

                        case "TRIAXIAL":
                        case "TRIAX":
                            _Orientation = Analyst.Orientation.Triaxial;
                            break;

                        case "X":
                            _Orientation = Analyst.Orientation.X;
                            break;

                        case "Y":
                            _Orientation = Analyst.Orientation.Y;
                            break;

                        case "Z":
                            _Orientation = Analyst.Orientation.Z;
                            break;

                        default:
                            _Orientation = Analyst.Orientation.None;
                            break;
                    }
                    _OrientationLoaded = true;
                }
                return _Orientation;
            }
            set
            {
                switch (value)
                {
                    case Analyst.Orientation.None:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "None");
                        break;

                    case Analyst.Orientation.Horizontal:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "Horizontal");
                        break;

                    case Analyst.Orientation.Vertical:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "Vertical");
                        break;

                    case Analyst.Orientation.Axial:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "Axial");
                        break;

                    case Analyst.Orientation.Radial:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "Radial");
                        break;

                    case Analyst.Orientation.Triaxial:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "Triax");
                        break;

                    case Analyst.Orientation.X:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "X");
                        break;

                    case Analyst.Orientation.Y:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "Y");
                        break;

                    case Analyst.Orientation.Z:
                        _OrientationLoaded = SetValueString("SKFCM_ASPF_Orientation", "Z");
                        break;

                    default:
                        _OrientationLoaded = false;
                        break;
                }
                if (_OrientationLoaded) _Orientation = value;
            }
        }
        public string OrientationPrefix
        {
            get
            {
                switch (Orientation)
                {
                    case Analyst.Orientation.None: return string.Empty;
                    case Analyst.Orientation.Horizontal: return "H";
                    case Analyst.Orientation.Vertical: return "V";
                    case Analyst.Orientation.Axial: return "A";
                    case Analyst.Orientation.Radial: return "R";
                    case Analyst.Orientation.Triaxial: return "T";
                    case Analyst.Orientation.X: return "X";
                    case Analyst.Orientation.Y: return "Y";
                    case Analyst.Orientation.Z: return "Z";
                    default: return string.Empty;
                }
            }
        }

        private string _OrientationPrefixFromName = string.Empty;
        public string OrientationPrefixFromName
        {
            get
            {
                if (string.IsNullOrEmpty(_OrientationPrefixFromName))
                {

                    //Obtem tecnica caso o DAD seja MI ou OS
                    string sDBStandard_Tecn = ((DadGroup == Analyst.DadGroup.MicrologAnalyzer) || (DadGroup == Analyst.DadGroup.Multilog) ? sDBStandard_Tecn = Tech : string.Empty);

                    //Detecta a direcao de medicao
                    string sDBStandard_TreeElem_NameTMP = NameRef();

                    if (sDBStandard_Tecn == string.Empty)
                    {
                        return string.Empty;
                    }
                    else if (sDBStandard_Tecn == "T")
                    {
                        _OrientationPrefixFromName = "T";
                    }
                    else if ((sDBStandard_TreeElem_NameTMP.IndexOf("H" + sDBStandard_Tecn.Substring(0, 1)) > 0) | (sDBStandard_TreeElem_NameTMP.IndexOf("H " + sDBStandard_Tecn.Substring(0, 1)) > 0))
                    {
                        _OrientationPrefixFromName = "H";
                    }
                    else if ((sDBStandard_TreeElem_NameTMP.IndexOf("V" + sDBStandard_Tecn.Substring(0, 1)) > 0) | (sDBStandard_TreeElem_NameTMP.IndexOf("V " + sDBStandard_Tecn.Substring(0, 1)) > 0))
                    {
                        _OrientationPrefixFromName = "V";
                    }
                    else if ((sDBStandard_TreeElem_NameTMP.IndexOf("A" + sDBStandard_Tecn.Substring(0, 1)) > 0) | (sDBStandard_TreeElem_NameTMP.IndexOf("A " + sDBStandard_Tecn.Substring(0, 1)) > 0))
                    {
                        _OrientationPrefixFromName = "A";
                    }
                    else if ((sDBStandard_TreeElem_NameTMP.IndexOf("X" + sDBStandard_Tecn.Substring(0, 1)) > 0) | (sDBStandard_TreeElem_NameTMP.IndexOf("X " + sDBStandard_Tecn.Substring(0, 1)) > 0))
                    {
                        _OrientationPrefixFromName = "X";
                    }
                    else if ((sDBStandard_TreeElem_NameTMP.IndexOf("Y" + sDBStandard_Tecn.Substring(0, 1)) > 0) | (sDBStandard_TreeElem_NameTMP.IndexOf("Y " + sDBStandard_Tecn.Substring(0, 1)) > 0))
                    {
                        _OrientationPrefixFromName = "Y";
                    }
                    else if ((sDBStandard_TreeElem_NameTMP.IndexOf("R" + sDBStandard_Tecn.Substring(0, 1)) > 0) | (sDBStandard_TreeElem_NameTMP.IndexOf("R " + sDBStandard_Tecn.Substring(0, 1)) > 0))
                    {
                        _OrientationPrefixFromName = "R";
                    }
                }
                return _OrientationPrefixFromName;
            }
        }

        private Point _ConditionalPoint = null;
        public Point ConditionalPoint
        {
            get
            {
                if (_ConditionalPoint == null)
                {
                    uint ConditionalPointId = ValueUInt("SKFCM_ASPF_Conditional_Point");
                    if (ConditionalPointId > 0) _ConditionalPoint = new Analyst.Point(Connection, ConditionalPointId);
                }
                return _ConditionalPoint;
            }
        }

        private Point _SpeedReferencePoint = null;
        public Point SpeedReferencePoint
        {
            get
            {
                if (_SpeedReferencePoint == null)
                {
                    uint SpeedReferencePointId = ValueUInt("SKFCM_ASPF_Speed_Reference_Id");
                    if (SpeedReferencePointId > 0) _SpeedReferencePoint = new Analyst.Point(Connection, SpeedReferencePointId);
                }
                return _SpeedReferencePoint;
            }
        }


        private static string GetNameRef_Last_Result = string.Empty;
        private string NameRef()
        {
            if (!string.IsNullOrEmpty(GetNameRef_Last_Result)) return GetNameRef_Last_Result;

            string sGetNameRef_NameRef = _TreeElem.Name;

            sGetNameRef_NameRef = sGetNameRef_NameRef.Trim();
            sGetNameRef_NameRef = sGetNameRef_NameRef.ToUpper();
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LA H", "01H");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LA V", "01V");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LA A", "01A");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LA R", "01R");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LC H", "02H");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LC V", "02V");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LC A", "02A");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LC R", "02R");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LO H", "02H");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LO V", "02V");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LO A", "02A");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("LO R", "02R");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("VEL H", "HV");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("VEL V", "VV");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("VEL A", "AV");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("ACC H", "HA");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("ACC V", "VA");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("ACC A", "AA");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("ENV H", "HE");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("ENV V", "VE");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("ENV A", "AE");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace(" NDE ", " ");
            sGetNameRef_NameRef = sGetNameRef_NameRef.Replace(" DE ", " ");

            if (Tech.StartsWith("E")) sGetNameRef_NameRef = sGetNameRef_NameRef.Replace("GE", "E");

            GetNameRef_Last_Result = sGetNameRef_NameRef.Trim();
            //while (GetNameRef_Last_Result.IndexOf(" ")>=0) GetNameRef_Last_Result = GetNameRef_Last_Result.Replace("  "," ");
            return GetNameRef_Last_Result;
        }

        public void Archive()
        {
            GenericTools.DebugMsg("Archive(): Starting");

            try
            {
                if (IsLoaded)
                {
                    DateTime tTemp;
                    string tTempDtg;

                    switch (ValueUInt("SKFCM_ASPF_SCHEDULE_Keep_Data_Unit"))
                    {
                        case 108: // Keep forever
                            GenericTools.DebugMsg("Archive(" + PointId.ToString() + "): Keep data forever...");
                            break;

                        case 109: // number of measurements
                            GenericTools.DebugMsg("Archive(" + PointId.ToString() + "): Number of measurement based");
                            break;

                        default: // time limited
                            GenericTools.DebugMsg("Archive(" + PointId.ToString() + "): Time limited...");
                            DateTime tActualLastData = DateTime.Now.AddSeconds((-1) * ValueUInt("SKFCM_ASPF_SCHEDULE_Keep_Data"));
                            DateTime tShortLastData = tActualLastData.AddSeconds((-1) * ValueUInt("SKFCM_ASPF_SCHEDULE_Keep_SArchive"));
                            DateTime tLongLastData = tShortLastData.AddSeconds((-1) * ValueUInt("SKFCM_ASPF_SCHEDULE_Keep_LArchive"));

                            _TreeElem.Connection.SQLUpdate("Measurement", "Status", "3", "PointId=" + PointId.ToString() + " and Status=1 and DataDtg<'" + GenericTools.DateTime(tActualLastData) + "'");

                            if (ValueUInt("SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit") < 108)
                            {
                                Connection.SQLUpdate("Measurement", "Status", "4", "PointId=" + PointId.ToString() + " and Status=3 and DataDtg<'" + GenericTools.DateTime(tShortLastData) + "'");
                                if (ValueUInt("SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit") < 108) Connection.SQLExec("delete from " + _TreeElem.Connection.Owner + "Measurement where PointId=" + PointId.ToString() + " and Status=4 and DataDtg<'" + GenericTools.DateTime(tLongLastData) + "'");

                                // Long
                                tTempDtg = Connection.SQLtoString("min(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and Status=4");
                                if (!string.IsNullOrEmpty(tTempDtg))
                                {
                                    tLongLastData = GenericTools.StrToDateTime(tTempDtg);
                                    tTemp = tLongLastData.AddSeconds(ValueUInt("SKFCM_ASPF_SCHEDULE_Move_LArchive"));
                                    do
                                    {
                                        Connection.SQLExec("delete from " + Connection.Owner + "Measurement where PointId=" + PointId.ToString() + " and Status=4 and DataDtg>'" + GenericTools.DateTime(tLongLastData) + "' and DataDtg<'" + GenericTools.DateTime(tTemp) + "'");
                                        tTempDtg = Connection.SQLtoString("min(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and Status=4 and DataDtg>'" + GenericTools.DateTime(tLongLastData) + "'");
                                        if (!string.IsNullOrEmpty(tTempDtg))
                                        {
                                            tLongLastData = GenericTools.StrToDateTime(tTempDtg);
                                            tTemp = tLongLastData.AddSeconds(ValueUInt("SKFCM_ASPF_SCHEDULE_Move_LArchive"));
                                        }
                                    }
                                    while ((tLongLastData < tShortLastData) & (!string.IsNullOrEmpty(tTempDtg)));
                                }
                            }

                            // Short
                            tTempDtg = Connection.SQLtoString("min(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and Status=3");
                            if (!string.IsNullOrEmpty(tTempDtg))
                            {
                                tShortLastData = GenericTools.StrToDateTime(tTempDtg);
                                tTemp = tShortLastData.AddSeconds(ValueUInt("SKFCM_ASPF_SCHEDULE_Move_SArchive"));
                                do
                                {
                                    Connection.SQLExec("delete from " + Connection.Owner + "Measurement where PointId=" + PointId.ToString() + " and Status=3 and DataDtg>'" + GenericTools.DateTime(tShortLastData) + "' and DataDtg<'" + GenericTools.DateTime(tTemp) + "'");
                                    tTempDtg = Connection.SQLtoString("min(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and Status=3 and DataDtg>'" + GenericTools.DateTime(tShortLastData) + "'");
                                    if (!string.IsNullOrEmpty(tTempDtg))
                                    {
                                        tShortLastData = GenericTools.StrToDateTime(tTempDtg);
                                        tTemp = tShortLastData.AddSeconds(ValueUInt("SKFCM_ASPF_SCHEDULE_Move_SArchive"));
                                    }
                                }
                                while ((tShortLastData < tActualLastData) & (!string.IsNullOrEmpty(tTempDtg)));
                            }

                            Connection.SQLExec("delete from " + Connection.Owner + "MeasReading where PointId=" + PointId.ToString() + " and MeasId not in (select MeasId from " + Connection.Owner + "Measurement where PointId=" + PointId.ToString() + ")");
                            Connection.SQLExec("delete from " + Connection.Owner + "MeasAlarm where PointId=" + PointId.ToString() + " and MeasId not in (select MeasId from " + Connection.Owner + "Measurement where PointId=" + PointId.ToString() + ")");
                            Connection.SQLExec("delete from " + Connection.Owner + "MeasAlarm where PointId=" + PointId.ToString() + " and ReadingId not in (select ReadingId from " + Connection.Owner + "MeasReading where PointId=" + PointId.ToString() + ")");


                            /*
                            string sShortLastData = SQLtoString("max(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and Status=2");
                            if (string.IsNullOrEmpty(sShortLastData)) sShortLastData = "19000101000000";
                            DateTime tShortLastData = GenericTools.StrToDateTime(sShortLastData);
                            DateTime tShortNextData = tShortLastData.AddSeconds(SKFCM_ASPF_SCHEDULE_Move_SArchive);
                    
                            string sLongLastData = SQLtoString("max(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and Status=3");
                            if (string.IsNullOrEmpty(sLongLastData)) sLongLastData = "19000101000000";
                            DateTime tLongLastData = GenericTools.StrToDateTime(sLongLastData);
                            DateTime tLongNextData = tShortLastData.AddSeconds(SKFCM_ASPF_SCHEDULE_Move_LArchive);

                            // Removida para teste em produção SQLExec("delete from " + Owner + "Measurement where PointId=" + PointId.ToString() + " and Status=1 and DataDtg<'" + GenericTools.DateTime(tShortNextData) + "'");

                            tMeasurement = DataTable("MeasId, Status, DataDtg", "Measurement", "PointId=" + PointId.ToString() + " and Status=1 order by DataDtg");
                    */
                            break;
                    }
                    //tMeasurement = DataTable("MeasId, Status, DataDtg", "Measurement", "PointId=" + PointId.ToString() + " and Status=1 order by DataDtg desc");
                    //Int32 InitialCount = tMeasurement.Rows.Count;
                    //if (InitialCount > 0)
                    //{
                    //}
                }
                else
                {
                    GenericTools.DebugMsg("Archive(" + PointId.ToString() + "): Not connected to DB");
                }
            }
            catch (Exception ex)
            {
                GenericTools.GetError("Archive(" + PointId.ToString() + "): " + ex.Message);
            }
            GenericTools.DebugMsg("Archive(" + PointId.ToString() + "): Finished");
        }

        /// <summary>Store a overall value on SKF @ptitude Analyst database</summary>
        /// <param name="OverallValue">Scalar value do be stored</param>
        /// <returns>Measurement unique id (MeasId)
        /// <para>Returns 0 if measurement already exists on database</para>
        /// </returns>
        public uint StoreOverallData(float OverallValue) { return StoreOverallData(DateTime.Now, OverallValue); }
        /// <summary>Store a overall value on SKF @ptitude Analyst database</summary>
        /// <param name="TimeStamp">Measurement time stamp</param>
        /// <param name="OverallValue">Scalar value do be stored</param>
        /// <returns>Measurement unique id (MeasId)
        /// <para>Returns 0 if measurement already exists on database</para>
        /// </returns>
        public uint StoreOverallData(DateTime TimeStamp, float OverallValue) { return StoreOverallData(GenericTools.DateTime(TimeStamp), OverallValue); }
        /// <summary>Store a overall value on SKF @ptitude Analyst database</summary>
        /// <param name="DataDtg">Measurement time stamp (YYYYMMDDHHMISS)</param>
        /// <param name="OverallValue">Scalar value do be stored</param>
        /// <returns>Measurement unique id (MeasId)
        /// <para>Returns 0 if measurement already exists on database</para>
        /// </returns>
        public uint StoreOverallData(string DataDtg, float OverallValue)
        {
            GenericTools.DebugMsg("StoreOverallData(\"" + DataDtg + "\", " + OverallValue.ToString() + "): Starting");
            uint ReturnValue = 0;

            try
            {
                if (IsLoaded)
                {
                    uint LastMeasId = Connection.SQLtoUInt("min(MeasId)", "MeasDtsRead", "PointId=" + PointId.ToString() + " and DataDtg='" + DataDtg + "' and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and round(OverallValue,4)=round(" + OverallValue.ToString().Replace(",", ".") + ",4)");
                    if (LastMeasId > 0)
                    {
                        ReturnValue = 0;
                    }
                    else
                    {
                        List<TableColumn> MeasurementColumns = new List<TableColumn>();

                        MeasurementColumns.Add(new TableColumn("MeasId", 0));
                        MeasurementColumns.Add(new TableColumn("PointId", PointId));
                        MeasurementColumns.Add(new TableColumn("AlarmLevel", 0));
                        MeasurementColumns.Add(new TableColumn("Status", 1));
                        MeasurementColumns.Add(new TableColumn("Include", 1));
                        MeasurementColumns.Add(new TableColumn("OperName", string.Empty));
                        MeasurementColumns.Add(new TableColumn("SerialNo", "STB"));
                        MeasurementColumns.Add(new TableColumn("DataDtg", DataDtg));
                        MeasurementColumns.Add(new TableColumn("MeasurementType", 0));

                        uint MeasIdTMP = Convert.ToUInt32(Connection.SQLInsert("Measurement", MeasurementColumns, "MeasId", "MeasId_seq"));

                        if (MeasIdTMP > 0)
                        {

                            List<TableColumn> MeasReadingColumns = new List<TableColumn>();

                            MeasReadingColumns.Add(new TableColumn("ReadingId", 0));
                            MeasReadingColumns.Add(new TableColumn("MeasId", MeasIdTMP));
                            MeasReadingColumns.Add(new TableColumn("PointId", PointId));
                            MeasReadingColumns.Add(new TableColumn("ReadingType", Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall")));
                            MeasReadingColumns.Add(new TableColumn("Channel", 1));
                            MeasReadingColumns.Add(new TableColumn("OverallValue", OverallValue));
                            //MeasReadingColumns.Add(new TableColumn("ExDWordVal1", 0));
                            // MeasReadingColumns.Add(new TableColumn("ExDoubleVal1", 0.0));
                            //MeasReadingColumns.Add(new TableColumn("ExDoubleVal2", 0.0));
                            //MeasReadingColumns.Add(new TableColumn("ExDoubleVal3", 0.0));
                            //MeasReadingColumns.Add(new TableColumn("ReadingHeader", null));
                            //MeasReadingColumns.Add(new TableColumn("ReadingData", null));

                            uint ReadingIdTMP = Convert.ToUInt32(Connection.SQLInsert("MeasReading", MeasReadingColumns, "ReadingId", "MeasRdngId_seq"));
                            if (ReadingIdTMP > 0)
                            {
                                Connection.SQLExec("delete from " + Connection.Owner + "MeasAlarm where PointId=" + PointId.ToString() + " and AlarmType=" + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall").ToString() + " and TableName='ScalarAlarm'");

                                List<TableColumn> MeasAlarmColumns = new List<TableColumn>();

                                MeasAlarmColumns.Add(new TableColumn("MeasAlarmId", 0));
                                MeasAlarmColumns.Add(new TableColumn("PointId", PointId));
                                MeasAlarmColumns.Add(new TableColumn("ReadingId", ReadingIdTMP));
                                MeasAlarmColumns.Add(new TableColumn("AlarmId", ScalarAlarmId));
                                MeasAlarmColumns.Add(new TableColumn("AlarmType", Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall")));
                                MeasAlarmColumns.Add(new TableColumn("AlarmLevel", GetScalarAlarmLevel(OverallValue)[0]));
                                MeasAlarmColumns.Add(new TableColumn("ResultType", GetScalarAlarmLevel(OverallValue)[1]));
                                MeasAlarmColumns.Add(new TableColumn("TableName", "ScalarAlarm"));
                                MeasAlarmColumns.Add(new TableColumn("MeasId", MeasIdTMP));
                                MeasAlarmColumns.Add(new TableColumn("Channel", 1));

                                uint MeasAlarmIdTMP = Convert.ToUInt32(Connection.SQLInsert("MeasAlarm", MeasAlarmColumns, "MeasAlarmId", "MeasAlarmId_seq"));

                                if (MeasAlarmIdTMP > 0) ReturnValue = MeasIdTMP;
                            }
                            if (ReturnValue != MeasIdTMP)
                            {
                                Connection.SQLExec("delete from " + Connection.Owner + "MeasReading where MeasId=" + MeasIdTMP.ToString());
                                Connection.SQLExec("delete from " + Connection.Owner + "Measurement where MeasId=" + MeasIdTMP.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("StoreOverallData(\"" + DataDtg + "\", " + OverallValue.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("StoreOverallData(\"" + DataDtg + "\", " + OverallValue.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }



        public uint StoreInspectionData(DateTime TimeStamp, int NumberOfItem) { return StoreInspectionData(GenericTools.DateTime(TimeStamp), NumberOfItem); }
        /// <summary>Store a inspection value on SKF @ptitude Analyst database</summary>
        /// <param name="DataDtg">Measurement time stamp (YYYYMMDDHHMISS)</param>
        /// <param name="OverallValue">Scalar value do be stored</param>
        /// <returns>Measurement unique id (MeasId)
        /// <para>Returns 0 if measurement already exists on database</para>
        /// </returns>
        private uint StoreInspectionData(string DataDtg, int OverallValue)
        {
            GenericTools.DebugMsg("StoreInspectionData(\"" + DataDtg + "\", " + OverallValue.ToString() + "): Starting");
            uint ReturnValue = 0;

            try
            {
                if (IsLoaded)
                {
                    uint LastMeasId = Connection.SQLtoUInt("min(MeasId)", "MeasDtsRead", "PointId=" + PointId.ToString() + " and DataDtg='" + DataDtg + "' and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection").ToString() + " and EXDWORDVAL1=" + OverallValue.ToString() + "");
                    if (LastMeasId > 0)
                    {
                        ReturnValue = 0;
                    }
                    else
                    {
                        List<TableColumn> MeasurementColumns = new List<TableColumn>();

                        MeasurementColumns.Add(new TableColumn("MeasId", 0));
                        MeasurementColumns.Add(new TableColumn("PointId", PointId));
                        MeasurementColumns.Add(new TableColumn("AlarmLevel", 0));
                        MeasurementColumns.Add(new TableColumn("Status", 1));
                        MeasurementColumns.Add(new TableColumn("Include", 1));
                        MeasurementColumns.Add(new TableColumn("OperName", string.Empty));
                        MeasurementColumns.Add(new TableColumn("SerialNo", "STB"));
                        MeasurementColumns.Add(new TableColumn("DataDtg", DataDtg));
                        MeasurementColumns.Add(new TableColumn("MeasurementType", 0));

                        uint MeasIdTMP = Convert.ToUInt32(Connection.SQLInsert("Measurement", MeasurementColumns, "MeasId", "MeasId_seq"));

                        if (MeasIdTMP > 0)
                        {

                            List<TableColumn> MeasReadingColumns = new List<TableColumn>();

                            MeasReadingColumns.Add(new TableColumn("ReadingId", 0));
                            MeasReadingColumns.Add(new TableColumn("MeasId", MeasIdTMP));
                            MeasReadingColumns.Add(new TableColumn("PointId", PointId));
                            MeasReadingColumns.Add(new TableColumn("ReadingType", Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection")));
                            MeasReadingColumns.Add(new TableColumn("Channel", 1));
                            MeasReadingColumns.Add(new TableColumn("OverallValue", 0));
                            MeasReadingColumns.Add(new TableColumn("EXDWORDVAL1", OverallValue));
                            MeasReadingColumns.Add(new TableColumn("EXDOUBLEVAL1", 0));
                            MeasReadingColumns.Add(new TableColumn("EXDOUBLEVAL2", 0));
                            MeasReadingColumns.Add(new TableColumn("EXDOUBLEVAL3", 0));

                            uint ReadingIdTMP = Convert.ToUInt32(Connection.SQLInsert("MeasReading", MeasReadingColumns, "ReadingId", "MeasRdngId_seq"));
                            if (ReadingIdTMP > 0)
                            {
                                Connection.SQLExec("delete from " + Connection.Owner + "MeasAlarm where PointId=" + PointId.ToString() + " and AlarmType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection").ToString() + " and TableName='ScalarAlarm'");

                                List<TableColumn> MeasAlarmColumns = new List<TableColumn>();

                                MeasAlarmColumns.Add(new TableColumn("MeasAlarmId", 0));
                                MeasAlarmColumns.Add(new TableColumn("PointId", PointId));
                                MeasAlarmColumns.Add(new TableColumn("ReadingId", ReadingIdTMP));
                                MeasAlarmColumns.Add(new TableColumn("AlarmId", ScalarAlarmId));
                                MeasAlarmColumns.Add(new TableColumn("AlarmType", Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection")));
                                MeasAlarmColumns.Add(new TableColumn("AlarmLevel", GetScalarAlarmLevel(OverallValue)[0]));
                                MeasAlarmColumns.Add(new TableColumn("ResultType", GetScalarAlarmLevel(OverallValue)[1]));
                                MeasAlarmColumns.Add(new TableColumn("TableName", "ScalarAlarm"));
                                MeasAlarmColumns.Add(new TableColumn("MeasId", MeasIdTMP));
                                MeasAlarmColumns.Add(new TableColumn("Channel", 1));

                                uint MeasAlarmIdTMP = Convert.ToUInt32(Connection.SQLInsert("MeasAlarm", MeasAlarmColumns, "MeasAlarmId", "MeasAlarmId_seq"));

                                if (MeasAlarmIdTMP > 0) ReturnValue = MeasIdTMP;
                            }
                            if (ReturnValue != MeasIdTMP)
                            {
                                Connection.SQLExec("delete from " + Connection.Owner + "MeasReading where MeasId=" + MeasIdTMP.ToString());
                                Connection.SQLExec("delete from " + Connection.Owner + "Measurement where MeasId=" + MeasIdTMP.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("StoreInspectionData(\"" + DataDtg + "\", " + OverallValue.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("StoreInspectionData(\"" + DataDtg + "\", " + OverallValue.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }



        private uint _ScalarAlarmId = 0;
        public uint ScalarAlarmId
        {
            get
            {
                if ((_ScalarAlarmId == 0) & IsLoaded)
                {
                    //Danilo 29/08/2016 - Cálculo de pontos derivados DS2SAM
                    _ScalarAlarmId = Connection.SQLtoUInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString() + " and Type=" + Registration.RegistrationId(Connection, "SKFCM_ASAT_Overall").ToString() + " and AlarmIndex=0");
                    if (_ScalarAlarmId < 1) _ScalarAlarmId = Connection.SQLtoUInt("ScalarAlrmId", "ScalarAlarm", "ElementId=" + PointId.ToString());
                }
                return _ScalarAlarmId;
            }
        }

        /// <summary>Get alarm level for last value</summary>
        /// <returns>Alarm level:
        /// <para>0: No alarm info</para>
        /// <para>1: Out of alarm</para>
        /// <para>2: Alert</para>
        /// <para>3: Danger</para>
        /// </returns>
        public int[] GetScalarAlarmLevel() { return GetScalarAlarmLevel(Connection.SQLtoFloat("OverallValue", "MeasDtsRead", "PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " order by DataDtg desc")); }
        //public int GetScalarAlarmLevel(Int32 PointId, Int32 MeasId) { return GetScalarAlarmLevel(PointId, SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + GetRegistrationId("SKFCM_ASMD_Overall")));}
        /// <summary>Get alarm level for specific value in a given point</summary>
        /// <param name="ReadingId">Reading unique id</param>
        /// <returns>Alarm level:
        /// <para>0: No alarm info</para>
        /// <para>1: Out of alarm</para>
        /// <para>2: Alert</para>
        /// <para>3: Danger</para></returns>
        public int[] GetScalarAlarmLevel(uint ReadingId) { return GetScalarAlarmLevel(Connection.SQLtoFloat("OverallValue", "MeasReading", "ReadingId=" + ReadingId.ToString())); }
        /// <summary>Get alarm level for specific value in a given point</summary>
        /// <param name="OverallValue">Overall value measured</param>
        /// <returns>Alarm level:
        /// <para>0: No alarm info</para>
        /// <para>1: Out of alarm</para>
        /// <para>2: Alert</para>
        /// <para>3: Danger</para></returns>
        public int[] GetScalarAlarmLevel(float OverallValue)
        {
            GenericTools.DebugMsg("GetScalarAlarmLevel(" + PointId.ToString() + ", " + OverallValue.ToString() + "): Starting");
            int[] ReturnValue = new int[2] { 2, 0 };

            try
            {
                if (IsLoaded)
                {
                    bool EnableDangerHi = false;
                    bool EnableDangerLo = false;
                    bool EnableAlertHi = false;
                    bool EnableAlertLo = false;
                    int AlarmMethod = 0;
                    float DangerLo = 0;
                    float DangerHi = 0;
                    float AlertLo = 0;
                    float AlertHi = 0;

                    DataTable ScalarAlarm = Connection.DataTable("ScalarAlrmId, ElementId, AlarmSet, EnableDangerHi, EnableDangerLo, EnableAlertHi, EnableAlertLo, AlarmMethod, DangerLo, DangerHi, AlertLo, AlertHi", "ScalarAlarm", "ScalarAlrmId=" + ScalarAlarmId.ToString());
                    if (ScalarAlarm.Rows.Count > 0)
                    {
                        EnableDangerHi = (Convert.ToInt16(ScalarAlarm.Rows[0]["EnableDangerHi"]) == 1);
                        EnableDangerLo = (Convert.ToInt16(ScalarAlarm.Rows[0]["EnableDangerLo"]) == 1);
                        EnableAlertHi = (Convert.ToInt16(ScalarAlarm.Rows[0]["EnableAlertHi"]) == 1);
                        EnableAlertLo = (Convert.ToInt16(ScalarAlarm.Rows[0]["EnableAlertLo"]) == 1);
                        AlarmMethod = Convert.ToInt16(ScalarAlarm.Rows[0]["AlarmMethod"]);
                        DangerLo = (float)Convert.ToDouble(ScalarAlarm.Rows[0]["DangerLo"]);
                        DangerHi = (float)Convert.ToDouble(ScalarAlarm.Rows[0]["DangerHi"]);
                        AlertLo = (float)Convert.ToDouble(ScalarAlarm.Rows[0]["AlertLo"]);
                        AlertHi = (float)Convert.ToDouble(ScalarAlarm.Rows[0]["AlertHi"]);

                        switch (AlarmMethod)
                        {
                            case 1: //Level
                                if (EnableAlertHi & (OverallValue >= AlertHi)) ReturnValue = new int[2] { 3, 1 };
                                if (EnableDangerHi & (OverallValue >= DangerHi)) ReturnValue = new int[2] { 4, 3 };
                                break;

                            case 2: //Window
                                if (EnableDangerLo & (OverallValue >= DangerLo)) ReturnValue = new int[2] { 4, 4 };
                                if (EnableDangerHi & (OverallValue <= DangerHi)) ReturnValue = new int[2] { 4, 3 };
                                if (EnableAlertLo & (OverallValue >= AlertLo)) ReturnValue = new int[2] { 3, 2 };
                                if (EnableAlertHi & (OverallValue <= AlertHi)) ReturnValue = new int[2] { 3, 1 };
                                break;

                            case 3: //Out of window
                                if (EnableAlertHi & (OverallValue >= AlertHi)) ReturnValue = new int[2] { 3, 1 };
                                if (EnableAlertLo & (OverallValue <= AlertLo)) ReturnValue = new int[2] { 3, 2 };
                                if (EnableDangerHi & (OverallValue >= DangerHi)) ReturnValue = new int[2] { 4, 3 };
                                if (EnableDangerLo & (OverallValue <= DangerLo)) ReturnValue = new int[2] { 4, 4 };
                                break;

                            default: //None
                                ReturnValue = new int[2] { 2, 0 };
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("GetScalarAlarmLevel(" + PointId.ToString() + ", " + OverallValue.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("GetScalarAlarmLevel(" + PointId.ToString() + ", " + OverallValue.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        private string _AlarmSet = string.Empty;
        public string AlarmSet
        {
            get
            {
                if (_AlarmSet == string.Empty)
                    _AlarmSet = Connection.SQLtoString("AlarmSet", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId);

                return _AlarmSet;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLUpdate("ScalarAlarm", "AlarmSet", value, "SCALARALRMID=" + this.ScalarAlarmId);
                    _AlarmSet = value;
                }
            }
        }

        private bool _EnableDangerHi = false;
        public bool EnableDangerHi
        {
            get
            {
                if (_EnableDangerHi == false)
                    _EnableDangerHi = ((Connection.SQLtoInt("ENABLEDANGERHI", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId) == 1) ? true : false);

                return _EnableDangerHi;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLExec("delete from " + Connection.Owner + "ScalarAlarm where SCALARALRMID!=" + this.ScalarAlarmId + " and ElementId=" + this.TreeElemId);
                    Connection.SQLUpdate("ScalarAlarm", "ENABLEDANGERHI", (value == true) ? 1 : 0, "SCALARALRMID=" + this.ScalarAlarmId);
                    _EnableDangerHi = value;
                    uint AlarmLevelTmp;
                    DataTable MeasAlarmTable = Connection.DataTable("*", "MeasAlarm", "TableName='ScalarAlarm' and PointId=" + TreeElemId.ToString());
                    foreach (DataRow MeasAlarmItem in MeasAlarmTable.Rows)
                    {
                        uint AlarmLevelOrig = Convert.ToUInt32(MeasAlarmItem["AlarmLevel"]);
                        AlarmLevelTmp = 0;
                        float OverallReadingTmp = (new MeasReadingOverall(Connection, Convert.ToUInt32(MeasAlarmItem["MeasId"]))).OverallValue;
                        if (((EnableDangerHi) && (OverallReadingTmp >= DangerHi)) || ((EnableDangerLo) && (OverallReadingTmp <= DangerLo)))
                            AlarmLevelTmp = 4;
                        else if (((EnableAlertHi) && (OverallReadingTmp >= AlertHi)) || ((EnableAlertLo) && (OverallReadingTmp <= AlertLo)))
                            AlarmLevelTmp = 3;
                        else if (EnableAlertHi || EnableAlertLo || EnableDangerHi || EnableDangerLo)
                            //if (OverallReadingTmp >= _DangerHi)
                            AlarmLevelTmp = 2;

                        if (AlarmLevelTmp != AlarmLevelOrig)
                        {
                            Connection.SQLUpdate("MeasAlarm", "AlarmLevel", AlarmLevelTmp, "TableName='ScalarAlarm' and AlarmLevel!=" + AlarmLevelTmp.ToString() + " and PointId=" + TreeElemId.ToString());
                            TreeElem.CalcAlarm();
                        }
                    }
                }
            }
        }
        private bool _EnableDangerLo = false;
        public bool EnableDangerLo
        {
            get
            {
                if (_EnableDangerLo == false)
                    _EnableDangerLo = ((Connection.SQLtoInt("ENABLEDANGERLO", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId) == 1) ? true : false);

                return _EnableDangerLo;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLExec("delete from " + Connection.Owner + "ScalarAlarm where SCALARALRMID!=" + this.ScalarAlarmId + " and ElementId=" + this.TreeElemId);
                    Connection.SQLUpdate("ScalarAlarm", "ENABLEDANGERLO", (value == true) ? 1 : 0, "SCALARALRMID=" + this.ScalarAlarmId);
                    _EnableDangerLo = value;
                    uint AlarmLevelTmp;
                    DataTable MeasAlarmTable = Connection.DataTable("*", "MeasAlarm", "TableName='ScalarAlarm' and PointId=" + TreeElemId.ToString());
                    foreach (DataRow MeasAlarmItem in MeasAlarmTable.Rows)
                    {
                        uint AlarmLevelOrig = Convert.ToUInt32(MeasAlarmItem["AlarmLevel"]);
                        AlarmLevelTmp = 0;
                        float OverallReadingTmp = (new MeasReadingOverall(Connection, Convert.ToUInt32(MeasAlarmItem["MeasId"]))).OverallValue;
                        if (((EnableDangerHi) && (OverallReadingTmp >= DangerHi)) || ((EnableDangerLo) && (OverallReadingTmp <= DangerLo)))
                            AlarmLevelTmp = 4;
                        else if (((EnableAlertHi) && (OverallReadingTmp >= AlertHi)) || ((EnableAlertLo) && (OverallReadingTmp <= AlertLo)))
                            AlarmLevelTmp = 3;
                        else if (EnableAlertHi || EnableAlertLo || EnableDangerHi || EnableDangerLo)
                            //if (OverallReadingTmp >= _DangerHi)
                            AlarmLevelTmp = 2;

                        if (AlarmLevelTmp != AlarmLevelOrig)
                        {
                            Connection.SQLUpdate("MeasAlarm", "AlarmLevel", AlarmLevelTmp, "TableName='ScalarAlarm' and AlarmLevel!=" + AlarmLevelTmp.ToString() + " and PointId=" + TreeElemId.ToString());
                            TreeElem.CalcAlarm();
                        }
                    }
                }
            }
        }
        private bool _EnableAlertHi = false;
        public bool EnableAlertHi
        {
            get
            {
                if (_EnableAlertHi == false)
                    _EnableAlertHi = ((Connection.SQLtoInt("ENABLEAlertHI", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId) == 1) ? true : false);

                return _EnableAlertHi;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLExec("delete from " + Connection.Owner + "ScalarAlarm where SCALARALRMID!=" + this.ScalarAlarmId + " and ElementId=" + this.TreeElemId);
                    Connection.SQLUpdate("ScalarAlarm", "ENABLEAlertHI", (value == true) ? 1 : 0, "SCALARALRMID=" + this.ScalarAlarmId);
                    _EnableAlertHi = value;
                    uint AlarmLevelTmp;
                    DataTable MeasAlarmTable = Connection.DataTable("*", "MeasAlarm", "TableName='ScalarAlarm' and PointId=" + TreeElemId.ToString());
                    foreach (DataRow MeasAlarmItem in MeasAlarmTable.Rows)
                    {
                        uint AlarmLevelOrig = Convert.ToUInt32(MeasAlarmItem["AlarmLevel"]);
                        AlarmLevelTmp = 0;
                        float OverallReadingTmp = (new MeasReadingOverall(Connection, Convert.ToUInt32(MeasAlarmItem["MeasId"]))).OverallValue;
                        if (((EnableDangerHi) && (OverallReadingTmp >= DangerHi)) || ((EnableDangerLo) && (OverallReadingTmp <= DangerLo)))
                            AlarmLevelTmp = 4;
                        else if (((EnableAlertHi) && (OverallReadingTmp >= AlertHi)) || ((EnableAlertLo) && (OverallReadingTmp <= AlertLo)))
                            AlarmLevelTmp = 3;
                        else if (EnableAlertHi || EnableAlertLo || EnableDangerHi || EnableDangerLo)
                            //if (OverallReadingTmp >= _DangerHi)
                            AlarmLevelTmp = 2;

                        if (AlarmLevelTmp != AlarmLevelOrig)
                        {
                            Connection.SQLUpdate("MeasAlarm", "AlarmLevel", AlarmLevelTmp, "TableName='ScalarAlarm' and AlarmLevel!=" + AlarmLevelTmp.ToString() + " and PointId=" + TreeElemId.ToString());
                            TreeElem.CalcAlarm();
                        }
                    }
                }
            }
        }
        private bool _EnableAlertLo = false;
        public bool EnableAlertLo
        {
            get
            {
                if (_EnableAlertLo == false)
                    _EnableAlertLo = ((Connection.SQLtoInt("ENABLEAlertLO", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId) == 1) ? true : false);

                return _EnableAlertLo;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLExec("delete from " + Connection.Owner + "ScalarAlarm where SCALARALRMID!=" + this.ScalarAlarmId + " and ElementId=" + this.TreeElemId);
                    Connection.SQLUpdate("ScalarAlarm", "ENABLEAlertLO", (value == true) ? 1 : 0, "SCALARALRMID=" + this.ScalarAlarmId);
                    _EnableAlertLo = value;
                    uint AlarmLevelTmp;
                    DataTable MeasAlarmTable = Connection.DataTable("*", "MeasAlarm", "TableName='ScalarAlarm' and PointId=" + TreeElemId.ToString());
                    foreach (DataRow MeasAlarmItem in MeasAlarmTable.Rows)
                    {
                        uint AlarmLevelOrig = Convert.ToUInt32(MeasAlarmItem["AlarmLevel"]);
                        AlarmLevelTmp = 0;
                        float OverallReadingTmp = (new MeasReadingOverall(Connection, Convert.ToUInt32(MeasAlarmItem["MeasId"]))).OverallValue;
                        if (((EnableDangerHi) && (OverallReadingTmp >= DangerHi)) || ((EnableDangerLo) && (OverallReadingTmp <= DangerLo)))
                            AlarmLevelTmp = 4;
                        else if (((EnableAlertHi) && (OverallReadingTmp >= AlertHi)) || ((EnableAlertLo) && (OverallReadingTmp <= AlertLo)))
                            AlarmLevelTmp = 3;
                        else if (EnableAlertHi || EnableAlertLo || EnableDangerHi || EnableDangerLo)
                            //if (OverallReadingTmp >= _DangerHi)
                            AlarmLevelTmp = 2;

                        if (AlarmLevelTmp != AlarmLevelOrig)
                        {
                            Connection.SQLUpdate("MeasAlarm", "AlarmLevel", AlarmLevelTmp, "TableName='ScalarAlarm' and AlarmLevel!=" + AlarmLevelTmp.ToString() + " and PointId=" + TreeElemId.ToString());
                            TreeElem.CalcAlarm();
                        }
                    }
                }
            }
        }
        private float _DangerHi = float.NaN;
        public float DangerHi
        {
            get
            {
                if (float.IsNaN(_DangerHi))
                    _DangerHi = Connection.SQLtoFloat("DANGERHI", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId);

                return _DangerHi;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLExec("delete from " + Connection.Owner + "ScalarAlarm where SCALARALRMID!=" + this.ScalarAlarmId + " and ElementId=" + this.TreeElemId);
                    Connection.SQLUpdate("ScalarAlarm", "DANGERHI", value, "SCALARALRMID=" + this.ScalarAlarmId);
                    _DangerHi = value;
                    uint AlarmLevelTmp;
                    DataTable MeasAlarmTable = Connection.DataTable("*", "MeasAlarm", "TableName='ScalarAlarm' and PointId=" + TreeElemId.ToString());
                    foreach (DataRow MeasAlarmItem in MeasAlarmTable.Rows)
                    {
                        uint AlarmLevelOrig = Convert.ToUInt32(MeasAlarmItem["AlarmLevel"]);
                        AlarmLevelTmp = 0;
                        float OverallReadingTmp = (new MeasReadingOverall(Connection, Convert.ToUInt32(MeasAlarmItem["MeasId"]))).OverallValue;
                        if (((EnableDangerHi) && (OverallReadingTmp >= DangerHi)) || ((EnableDangerLo) && (OverallReadingTmp <= DangerLo)))
                            AlarmLevelTmp = 4;
                        else if (((EnableAlertHi) && (OverallReadingTmp >= AlertHi)) || ((EnableAlertLo) && (OverallReadingTmp <= AlertLo)))
                            AlarmLevelTmp = 3;
                        else if (EnableAlertHi || EnableAlertLo || EnableDangerHi || EnableDangerLo)
                            //if (OverallReadingTmp >= _DangerHi)
                            AlarmLevelTmp = 2;

                        if (AlarmLevelTmp != AlarmLevelOrig)
                        {
                            Connection.SQLUpdate("MeasAlarm", "AlarmLevel", AlarmLevelTmp, "TableName='ScalarAlarm' and AlarmLevel!=" + AlarmLevelTmp.ToString() + " and PointId=" + TreeElemId.ToString());
                            TreeElem.CalcAlarm();
                        }
                    }
                }
            }
        }
        private float _DangerLo = float.NaN;
        public float DangerLo
        {
            get
            {
                if (float.IsNaN(_DangerLo))
                    _DangerLo = Connection.SQLtoFloat("DangerLo", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId);

                return _DangerLo;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLExec("delete from " + Connection.Owner + "ScalarAlarm where SCALARALRMID!=" + this.ScalarAlarmId + " and ElementId=" + this.TreeElemId);
                    Connection.SQLUpdate("ScalarAlarm", "DangerLo", value, "SCALARALRMID=" + this.ScalarAlarmId);
                    _DangerLo = value;
                    uint AlarmLevelTmp;
                    DataTable MeasAlarmTable = Connection.DataTable("*", "MeasAlarm", "TableName='ScalarAlarm' and PointId=" + TreeElemId.ToString());
                    foreach (DataRow MeasAlarmItem in MeasAlarmTable.Rows)
                    {
                        uint AlarmLevelOrig = Convert.ToUInt32(MeasAlarmItem["AlarmLevel"]);
                        AlarmLevelTmp = 0;
                        float OverallReadingTmp = (new MeasReadingOverall(Connection, Convert.ToUInt32(MeasAlarmItem["MeasId"]))).OverallValue;
                        if (((EnableDangerHi) && (OverallReadingTmp >= DangerHi)) || ((EnableDangerLo) && (OverallReadingTmp <= DangerLo)))
                            AlarmLevelTmp = 4;
                        else if (((EnableAlertHi) && (OverallReadingTmp >= AlertHi)) || ((EnableAlertLo) && (OverallReadingTmp <= AlertLo)))
                            AlarmLevelTmp = 3;
                        else if (EnableAlertHi || EnableAlertLo || EnableDangerHi || EnableDangerLo)
                            //if (OverallReadingTmp >= _DangerHi)
                            AlarmLevelTmp = 2;

                        if (AlarmLevelTmp != AlarmLevelOrig)
                        {
                            Connection.SQLUpdate("MeasAlarm", "AlarmLevel", AlarmLevelTmp, "TableName='ScalarAlarm' and AlarmLevel!=" + AlarmLevelTmp.ToString() + " and PointId=" + TreeElemId.ToString());
                            TreeElem.CalcAlarm();
                        }
                    }
                }
            }
        }
        private float _AlertHi = float.NaN;
        public float AlertHi
        {
            get
            {
                if (float.IsNaN(_AlertHi))
                    _AlertHi = Connection.SQLtoFloat("AlertHI", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId);

                return _AlertHi;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLExec("delete from " + Connection.Owner + "ScalarAlarm where SCALARALRMID!=" + this.ScalarAlarmId + " and ElementId=" + this.TreeElemId);
                    Connection.SQLUpdate("ScalarAlarm", "AlertHI", value, "SCALARALRMID=" + this.ScalarAlarmId);
                    _AlertHi = value;
                    uint AlarmLevelTmp;
                    DataTable MeasAlarmTable = Connection.DataTable("*", "MeasAlarm", "TableName='ScalarAlarm' and PointId=" + TreeElemId.ToString());
                    foreach (DataRow MeasAlarmItem in MeasAlarmTable.Rows)
                    {
                        uint AlarmLevelOrig = Convert.ToUInt32(MeasAlarmItem["AlarmLevel"]);
                        AlarmLevelTmp = 0;
                        float OverallReadingTmp = (new MeasReadingOverall(Connection, Convert.ToUInt32(MeasAlarmItem["MeasId"]))).OverallValue;
                        if (((EnableDangerHi) && (OverallReadingTmp >= DangerHi)) || ((EnableDangerLo) && (OverallReadingTmp <= DangerLo)))
                            AlarmLevelTmp = 4;
                        else if (((EnableAlertHi) && (OverallReadingTmp >= AlertHi)) || ((EnableAlertLo) && (OverallReadingTmp <= AlertLo)))
                            AlarmLevelTmp = 3;
                        else if (EnableAlertHi || EnableAlertLo || EnableDangerHi || EnableDangerLo)
                            //if (OverallReadingTmp >= _DangerHi)
                            AlarmLevelTmp = 2;

                        if (AlarmLevelTmp != AlarmLevelOrig)
                        {
                            Connection.SQLUpdate("MeasAlarm", "AlarmLevel", AlarmLevelTmp, "TableName='ScalarAlarm' and AlarmLevel!=" + AlarmLevelTmp.ToString() + " and PointId=" + TreeElemId.ToString());
                            TreeElem.CalcAlarm();
                        }
                    }
                }
            }
        }
        private float _AlertLo = float.NaN;
        public float AlertLo
        {
            get
            {
                if (float.IsNaN(_AlertLo))
                    _AlertLo = Connection.SQLtoFloat("AlertLo", "ScalarAlarm", "SCALARALRMID=" + this.ScalarAlarmId);

                return _AlertLo;
            }
            set
            {
                if (ScalarAlarmId != 0)
                {
                    Connection.SQLExec("delete from " + Connection.Owner + "ScalarAlarm where SCALARALRMID!=" + this.ScalarAlarmId + " and ElementId=" + this.TreeElemId);
                    Connection.SQLUpdate("ScalarAlarm", "AlertLo", value, "SCALARALRMID=" + this.ScalarAlarmId);
                    _AlertLo = value;
                    uint AlarmLevelTmp;
                    DataTable MeasAlarmTable = Connection.DataTable("*", "MeasAlarm", "TableName='ScalarAlarm' and PointId=" + TreeElemId.ToString());
                    foreach (DataRow MeasAlarmItem in MeasAlarmTable.Rows)
                    {
                        uint AlarmLevelOrig = Convert.ToUInt32(MeasAlarmItem["AlarmLevel"]);
                        AlarmLevelTmp = 0;
                        float OverallReadingTmp = (new MeasReadingOverall(Connection, Convert.ToUInt32(MeasAlarmItem["MeasId"]))).OverallValue;
                        if (((EnableDangerHi) && (OverallReadingTmp >= DangerHi)) || ((EnableDangerLo) && (OverallReadingTmp <= DangerLo)))
                            AlarmLevelTmp = 4;
                        else if (((EnableAlertHi) && (OverallReadingTmp >= AlertHi)) || ((EnableAlertLo) && (OverallReadingTmp <= AlertLo)))
                            AlarmLevelTmp = 3;
                        else if (EnableAlertHi || EnableAlertLo || EnableDangerHi || EnableDangerLo)
                            //if (OverallReadingTmp >= _DangerHi)
                            AlarmLevelTmp = 2;

                        if (AlarmLevelTmp != AlarmLevelOrig)
                        {
                            Connection.SQLUpdate("MeasAlarm", "AlarmLevel", AlarmLevelTmp, "TableName='ScalarAlarm' and AlarmLevel!=" + AlarmLevelTmp.ToString() + " and PointId=" + TreeElemId.ToString());
                            TreeElem.CalcAlarm();
                        }
                    }
                }
            }
        }

        /** private string _PointType = string.Empty;
        public string PointType
        {
            get
            {
                if (string.IsNullOrEmpty(_PointType))
                    _PointType = Registration.Signature(Connection, Convert.ToUInt32(ValueString("SKFCM_ASPF_Point_Type_Id")));
                return _PointType;
            }
        }
        **/

        public DataTable getBearings()
        {
            DataTable _return = new DataTable();
            try
            {
                DataTable dt_Bear = new DataTable();
                dt_Bear = Connection.DataTable("select NAME from FREQENTRIES WHERE FSID IN (select FSID from FREQASSIGN WHERE ELEMENTID=" + TreeElemId + ") AND GEOMETRYTABLE='SKFCM_ASGI_BearingFreqType'");


                _return.Columns.Add("Bearing");
                _return.Columns.Add("Manufacture");
                _return.Columns.Add("BPFO");
                _return.Columns.Add("BPFI");
                _return.Columns.Add("BSF");
                _return.Columns.Add("FTF");

                if (dt_Bear.Rows.Count > 0)
                {
                    string name;

                    for (int i = 0; i < dt_Bear.Rows.Count; i++)
                    {
                        name = dt_Bear.Rows[i][0].ToString();
                        string[] bear = name.Split('(');

                        if (bear.Length > 1)
                        {
                            string[] manu = bear[1].Split(')');
                            DataTable temp = new DataTable();
                            temp = Connection.DataTable("select * from BEARING where NAME='" + bear[0].Trim() + "' and MANUFACTURE = '" + manu[0].Replace('(', ' ').Trim() + "'");
                            string[] RowInfo = { temp.Rows[0]["NAME"].ToString(), temp.Rows[0]["MANUFACTURE"].ToString(), temp.Rows[0]["BPFO"].ToString(), temp.Rows[0]["BPFI"].ToString(), temp.Rows[0]["BSF"].ToString(), temp.Rows[0]["FTF"].ToString() };
                            _return.Rows.Add(RowInfo);
                        }
                    }
                }

            }
            catch (Exception ex)
            {


            }
            return _return;
        }


        public float AlertHi_StaticticAlarm(float Average, float StdDev)
        {
            float Calc = float.NaN;

            try
            {
                Calc = Average + (StdDev);
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("AlertHi_StaticticAlarm: Error: " + ex.Message);
                Calc = float.NaN;
            }
            return Calc;
        }


        public bool DerivatedCalc()
        {
            bool ReturnValue = false;

            //Check if point type is DerivatedCalculated
            if (PointType == Analyst.PointType.DerivedPointCalculated)
            {
                uint ExpId = Connection.SQLtoUInt("ExpressionId", "DPExpressionDef", "OwnerId=" + TreeElemId.ToString());
                if (ExpId > 0)
                {
                    if (
                            (Connection.SQLtoUInt("count(*)", "DPExpressionFormula", "ItemKey=3009 and ItemType=1002 and ExpId=" + ExpId.ToString()) > 0)
                        & (Connection.SQLtoUInt("count(*)", "DPExpressionFormula", "ItemKey=2006 and ItemType=1001 and ExpId=" + ExpId.ToString()) > 0)
                        & (Connection.SQLtoUInt("count(*)", "DPExpressionFormula", "ItemKey=2007 and ItemType=1001 and ExpId=" + ExpId.ToString()) > 0)
                        )
                    {
                        uint ItemKey = Connection.SQLtoUInt("ItemKey", "DPExpressionFormula", "ItemType=1003 and ExpId=" + ExpId.ToString());
                        if (ItemKey > 0)
                        {
                            uint SourcePointId = Connection.SQLtoUInt("SourcePtId", "DPExpressionVarRef", "DPId=" + TreeElemId.ToString() + " and ExpId=" + ExpId.ToString() + " and VarKey=" + ItemKey.ToString());
                            if (SourcePointId > 0)
                            {
                                DataTable UnprocessedMeasurements = Connection.DataTable("*", "Measurement", "PointId=" + SourcePointId.ToString() + " and DataDtg>='" + GenericTools.DateTime(LastMeas.DataDTG) + "' order by DataDtg asc");
                                if (UnprocessedMeasurements.Rows.Count > 1)
                                {
                                    ReturnValue = true;
                                    Measurement UnprocessedMeasurement;
                                    Measurement PreviousMeasurement;
                                    float NewValue;
                                    for (int i = 1; i < UnprocessedMeasurements.Rows.Count; i++)
                                    {
                                        PreviousMeasurement = new Measurement(Connection, Convert.ToUInt32(UnprocessedMeasurements.Rows[i - 1]["MeasId"]));
                                        UnprocessedMeasurement = new Measurement(Connection, Convert.ToUInt32(UnprocessedMeasurements.Rows[i]["MeasId"]));
                                        NewValue = 100 * (UnprocessedMeasurement.OverallReading.OverallValue - PreviousMeasurement.OverallReading.OverallValue) / PreviousMeasurement.OverallReading.OverallValue;
                                        ReturnValue = (ReturnValue & (StoreOverallData(UnprocessedMeasurement.DataDTG, NewValue) > 0));
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return ReturnValue;
        }

        public bool DerivatedAverageCalc(int Average = 10)
        {
            bool ReturnValue = false;

            if (PointType == Analyst.PointType.DerivedPointCalculated)
            {
                uint ExpId = Connection.SQLtoUInt("ExpressionId", "DPExpressionDef", "OwnerId=" + TreeElemId.ToString());
                if (ExpId > 0)
                {
                    uint SourcePointId = Connection.SQLtoUInt("SourcePtId", "DPExpressionVarRef", "DPId=" + TreeElemId.ToString() + " and ExpId=" + ExpId.ToString());// + " and VarKey=" + ItemKey.ToString());
                    if (SourcePointId > 0)
                    {
                        Measurement Measurement_Derived = new Measurement(new Point(Connection, TreeElemId));

                        DataTable UnprocessedMeasurements = Connection.DataTable("*", "Measurement", "PointId=" + SourcePointId.ToString() + " and DataDtg>'" + GenericTools.DateTime(Measurement_Derived.DataDTG) + "' order by DataDtg asc");
                        if (UnprocessedMeasurements.Rows.Count > 1)
                        {
                            ReturnValue = true;
                            Measurement UnprocessedMeasurement;

                            float NewValue;
                            float AverageMeasurement;
                            ///for (int i = 0; i < UnprocessedMeasurements.Rows.Count; i++)
                            foreach (DataRow dr in UnprocessedMeasurements.Rows)
                            {
                                StringBuilder SQL = new StringBuilder();
                                SQL.AppendLine(" SELECT AVG(MR.OVERALLVALUE) ");
                                SQL.AppendLine(" FROM  ");
                                SQL.AppendLine(" 	MEASREADING MR  ");
                                SQL.AppendLine(" WHERE ");
                                SQL.AppendLine(" 	MR.MEASID IN ( ");
                                SQL.AppendLine(" 		SELECT  ");
                                SQL.AppendLine(" 			TOP " + Average + " ME.MEASID  ");
                                SQL.AppendLine(" 		FROM  ");
                                SQL.AppendLine(" 			MEASUREMENT ME  ");
                                SQL.AppendLine(" 		wHERE  ");
                                SQL.AppendLine(" 			ME.POINTID=" + SourcePointId.ToString());
                                SQL.AppendLine(" 			and ME.DATADTG <'" + dr["DATADTG"].ToString() + "'  ");
                                SQL.AppendLine(" 		ORDER BY DATADTG DESC ");
                                SQL.AppendLine(" 	)  ");
                                SQL.AppendLine(" and MR.READINGTYPE='" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall") + "'");


                                UnprocessedMeasurement = new Measurement(Connection, Convert.ToUInt32(dr["MeasId"].ToString()));
                                AverageMeasurement = Connection.SQLtoFloat(SQL.ToString());

                                if (float.IsNaN(AverageMeasurement))
                                {
                                    AverageMeasurement = 0;//UnprocessedMeasurement.OverallReading.OverallValue;
                                    NewValue = 0;
                                }
                                else
                                {
                                    NewValue = ((UnprocessedMeasurement.OverallReading.OverallValue / AverageMeasurement) - 1) * 100;
                                }
                                //float NewValue_Original = 100 * (UnprocessedMeasurement.OverallReading.OverallValue - AverageMeasurement) / AverageMeasurement;

                                ReturnValue = (ReturnValue & (StoreOverallData(UnprocessedMeasurement.DataDTG, NewValue) > 0));
                            }
                        }
                    }
                }
            }
            return ReturnValue;
        }

        public float[] ScalarStats(int nMinCount, float nMinValidValue, float nMaxValidValue, int nNumberOfRead, float nMinChangeRatio)
        {
            GenericTools.DebugMsg("GetScalarStats(nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + ", nMinChangeRatio=" + nMinChangeRatio.ToString() + "): Starting...");

            float[] FirstReturnValue = new float[2];
            float[] SecondReturnValue = new float[2];
            float[] ReturnValue = new float[2];

            FirstReturnValue[0] = Connection.SQLtoFloat("avg(OverallValue)", "MeasReading", "PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and RowNum<=" + nMinCount.ToString() + " and (OverallValue between " + Convert.ToString(nMinValidValue) + " and " + Convert.ToString(nMaxValidValue) + ") order by ReadingId desc");
            FirstReturnValue[1] = Connection.SQLtoFloat("stddev(OverallValue)", "MeasReading", "PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and RowNum<=" + nMinCount.ToString() + " and (OverallValue between " + nMinValidValue.ToString() + " and " + nMaxValidValue.ToString() + ") order by ReadingId desc");

            if (nNumberOfRead <= 1)
                ReturnValue = FirstReturnValue;
            else
            {
                //if ((Math.Abs(nMinValidValue / (FirstReturnValue[0] - (2 * FirstReturnValue[1]))) >= nMinChangeRatio) | (Math.Abs(nMaxValidValue / (FirstReturnValue[0] + (2 * FirstReturnValue[1]))) >= nMinChangeRatio))
                //{
                //    ReturnValue = FirstReturnValue;
                //}
                //else
                //{
                SecondReturnValue = ScalarStats(nMinCount, (FirstReturnValue[0] - (2 * FirstReturnValue[1])), (FirstReturnValue[0] + (2 * FirstReturnValue[1])), nNumberOfRead - 1, nMinChangeRatio);

                if (float.IsNaN(SecondReturnValue[0]) | float.IsInfinity(SecondReturnValue[0]) | (SecondReturnValue[0] <= float.MinValue))
                    ReturnValue = FirstReturnValue;
                else
                    ReturnValue = SecondReturnValue;
                //}
            }

            GenericTools.DebugMsg("GetScalarStats(nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + ", nMinChangeRatio=" + nMinChangeRatio.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }

        ///<summary>Statistical Alarm Calculate</summary>
        ///<param name="NumberOfRead">OBSOLETE Number of measurements to use for calculate the statistical alarm</param>
        ///<param name="MinCount">Minimun Number of measurements to use for calculate the statistical alarm</param>
        ///<param name="DataRange">Datatime range of measurements to use for calculate the statistical alarm</param>
        ///<param name="RecordAlarms">Use TRUE to record in SCALARALARM the new calculated value</param>
        ///<returns>Return 4 Parameters - Average, StdDeviation, AlertHi, AlarmHi. if RecordAlarms==FALSE AlertHi and AlarmHi will be float.NAN </returns>
        public float[] StatisticAlarmCalc(uint NumberOfRead, uint MinCount, DateTime[,] DataRange, double Alta, double Baixa, bool RecordAlarms = false, bool storeValue = true, string Nome_Do_Alarme = "")
        {
            float[] ReturnValue = new float[4];

            GenericTools.DebugMsg("ScalarStdAverage(PointId=" + PointId.ToString() + " MinimunCounts=" + MinCount.ToString() + " NumberOfRead=" + NumberOfRead.ToString() + " Exclude_Dates.Rows= " + DataRange.Length.ToString());
            StringBuilder Dates = new StringBuilder();
            Dates.Clear();

            string Range = string.Empty;

            if (DataRange.Length > 0)
            {

                Dates.Append(" and MEASID in (select Me.MEASID from " + Connection.Owner + "Measurement Me where Me.PointId=" + PointId);
                //Dates.Append(" and Me.DataDtg>'20120000000000'"); // Especial para IPMG para excluir alteração de configuração")
                Dates.Append(" and ("); // Especial para IPMG para excluir alteração de configuração")


                for (int i = 0; i < (DataRange.Length / 2); i++)
                {
                    string inicio = string.Format("{0:yyyyMMddHHmmss}", DataRange[0, i]);
                    string fim = string.Format("{0:yyyyMMddHHmmss}", DataRange[1, i]);
                    Range = Range + "DE: " + inicio + " ATÉ: " + fim + " | ";

                    Dates.AppendFormat((i == 0 ? string.Empty : " OR") + " (Me.DataDtg between '{0}' and '{1}')", inicio, fim);
                }
                Dates.Append("))");
            }

            GenericTools.DebugMsg("Range: " + Range);

            DataTable StatsData = new DataTable();
            GenericTools.DebugMsg("Querying: " + Connection.DBType.ToString());

            string Select = "round(avg(OverallValue),4) as AverageVal, round(" + (Connection.DBType == DBType.Oracle ? "stddev" : "stdev") + "(OverallValue),4) as StdDevVal";
            string Where = "PointId=" + PointId.ToString() +
                                                    " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() +
                                                        Dates.ToString();

            if (Connection.SQLtoInt("count(*)", "MeasReading", Where) >= MinCount)
                StatsData = this.Connection.DataTable(Select, "MeasReading", Where);

            GenericTools.DebugMsg("Stats Count: " + StatsData.Rows.Count);
            if (StatsData.Rows.Count > 0)
            {
                GenericTools.DebugMsg("ReturnValue[0]: " + StatsData.Rows[0]["AverageVal"].ToString());
                ReturnValue[0] = float.Parse(StatsData.Rows[0]["AverageVal"].ToString());

                GenericTools.DebugMsg("ReturnValue[1]: " + StatsData.Rows[0]["StdDevVal"].ToString());
                ReturnValue[1] = float.Parse(StatsData.Rows[0]["StdDevVal"].ToString());

                float DangerHi_Temp = float.NaN;
                float AlertHi_Temp = float.NaN;

                if (RecordAlarms == true)
                {
                    Point pt = new Point(TreeElem);
                    float Average = ReturnValue[0];
                    float StdDev = ReturnValue[1];
                    float DangerHi_4Calc = Average + ((float)Alta * StdDev);
                    float DangerHi_StatCalc = pt.DangerHi_ScalarAlarmHigh(pt.TreeElemId) - StdDev;
                    float AlertHi = pt.AlertHi_StaticticAlarm(Average, (float)Baixa * StdDev);

                    GenericTools.DebugMsg("Parameters catched: " + Average + ", " + StdDev + ", " + DangerHi_4Calc + ", " + DangerHi_StatCalc + ", " + AlertHi);

                    if (!float.IsNaN(AlertHi))
                    {

                        AlertHi_Temp = AlertHi;
                        //pt.AlertHi = AlertHi;
                        if (Nome_Do_Alarme != "")
                            pt.AlarmSet = Nome_Do_Alarme;

                        //pt.AlarmSet = "";
                        GenericTools.WriteLog("Point: " + pt.Name + " AlertHi: " + AlertHi.ToString());

                    }
                    if (DangerHi_StatCalc >= DangerHi_4Calc)
                    {
                        if (!float.IsNaN(DangerHi_StatCalc))
                        {

                            DangerHi_Temp = DangerHi_StatCalc;

                            //pt.DangerHi = DangerHi_StatCalc;
                            //pt.AlarmSet = "";
                            if (Nome_Do_Alarme != "")
                                pt.AlarmSet = Nome_Do_Alarme;

                            GenericTools.WriteLog("Point: " + pt.Name + " DangerHi (Second Peak - StdDev): " + DangerHi_StatCalc.ToString());
                        }
                    }
                    else
                    {
                        if (!float.IsNaN(DangerHi_StatCalc))
                        {
                            DangerHi_Temp = DangerHi_4Calc;
                            //pt.DangerHi = DangerHi_4Calc;
                            //pt.AlarmSet = "";
                            if (Nome_Do_Alarme != "")
                                pt.AlarmSet = Nome_Do_Alarme;

                            GenericTools.WriteLog("Point: " + pt.Name + " DangerHi (Avg+(" + Alta + "*StdDev)): " + DangerHi_StatCalc.ToString());
                        }
                    }
                    if (storeValue)
                    {
                        if (!float.IsNaN(DangerHi_Temp)) pt.DangerHi = DangerHi_Temp;
                        if (!float.IsNaN(AlertHi_Temp)) pt.AlertHi = AlertHi_Temp;
                        if (!float.IsNaN(DangerHi_Temp) && !float.IsNaN(AlertHi_Temp)) pt.AlarmSet = "";
                    }

                    ReturnValue[2] = DangerHi_Temp;
                    ReturnValue[3] = AlertHi_Temp;
                }
                else
                {
                    ReturnValue[2] = float.NaN;
                    ReturnValue[3] = float.NaN;
                }
                GenericTools.DebugMsg("ScalarStdAverage Returnin, AVG:" + ReturnValue[0] + " and STDEV: " + ReturnValue[1]);
            }
            else
            {
                ReturnValue[0] = float.NaN;
                ReturnValue[1] = float.NaN;
                ReturnValue[2] = float.NaN;
                ReturnValue[3] = float.NaN;
            }


            return ReturnValue;
        }

        public float DangerHi_ScalarAlarmHigh(uint PointId)
        {
            float ReturnValue = float.NaN;

            GenericTools.DebugMsg("ScalarAlarmHighAndStd(PointId=" + PointId.ToString());

            if (Connection.SQLtoInt("COUNT(*)", "MeasReading", "PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString()) >= 2)
            {
                float SecPeakResult = float.NaN;
                float StdDevResult = float.NaN;

                DataTable dt_Overall = new DataTable();
                dt_Overall = Connection.DataTable("OVERALLVALUE", "MeasReading", "PointId=" + PointId.ToString() +
                                                  " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " ORDER BY OVERALLVALUE desc");

                SecPeakResult = (float)Convert.ToDouble(dt_Overall.Rows[1][0]);

                //if (Connection.DBType == DBType.Oracle)
                //    StdDevResult = Connection.SQLtoFloat("stddev(OverallValue)", "MeasReading"
                //                                , "PointId=" + PointId.ToString() +
                //                                " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString());

                //if (Connection.DBType == DBType.MSSQL)
                //    StdDevResult = Connection.SQLtoFloat("stdev(OverallValue)", "MeasReading"
                //                                , "PointId=" + PointId.ToString() +
                //                                " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString());

                ReturnValue = SecPeakResult;
            }

            return ReturnValue;
        }


        public float ScalarAverage(int nMinCount, float nMinValidValue, float nMaxValidValue, int nNumberOfRead, float nMinChangeRatio)
        {
            GenericTools.DebugMsg("GetScalarAverage(nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + ", nMinChangeRatio=" + nMinChangeRatio.ToString() + "): Starting...");

            float nFirstAverage = float.NaN;
            float nFirstStdDev = float.NaN;
            float nSecondAverage = float.NaN;
            float ReturnValue = float.NaN;

            DataTable StatsData = Connection.DataTable("avg(OverallValue) as AverageVal, stddev(OverallValue) as StdDevVal", "MeasReading", "PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and RowNum<=" + nMinCount.ToString() + " and (OverallValue between " + Convert.ToString(nMinValidValue) + " and " + Convert.ToString(nMaxValidValue) + ") order by ReadingId desc");
            if (StatsData.Rows.Count > 0)
            {
                nFirstAverage = (float)Convert.ToDouble(StatsData.Rows[0]["AverageVal"]); // SQLtoFloat("avg(OverallValue)", "MeasReading", "PointId=" + nPointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall") + " and RowNum<=" + nMinCount.ToString() + " and (OverallValue between " + Convert.ToString(nMinValidValue) + " and " + Convert.ToString(nMaxValidValue) + ") order by ReadingId desc");
                if (nNumberOfRead <= 1)
                    ReturnValue = nFirstAverage;
                else
                {
                    nFirstStdDev = (float)Convert.ToDouble(StatsData.Rows[0]["StdDevVal"]); // SQLtoFloat("stddev(OverallValue)", "MeasReading", "PointId=" + nPointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall") + " and RowNum<=" + nMinCount.ToString() + " and (OverallValue between " + nMinValidValue.ToString() + " and " + nMaxValidValue.ToString() + ") order by ReadingId desc");
                    //if ((Math.Abs(nMinValidValue / (nFirstAverage - (2 * nFirstStdDev))) >= nMinChangeRatio) | (Math.Abs(nMaxValidValue / (nFirstAverage + (2 * nFirstStdDev))) >= nMinChangeRatio)) return nFirstAverage;
                    nSecondAverage = ScalarAverage(nMinCount, (nFirstAverage - (2 * nFirstStdDev)), (nFirstAverage + (2 * nFirstStdDev)), nNumberOfRead - 1, nMinChangeRatio);
                    if (float.IsNaN(nSecondAverage) | float.IsInfinity(nSecondAverage) | (nSecondAverage <= float.MinValue))
                        ReturnValue = nFirstAverage;
                    else
                        ReturnValue = nSecondAverage;
                }
            }

            GenericTools.DebugMsg("GetScalarAverage(nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + ", nMinChangeRatio=" + nMinChangeRatio.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }
        public float ScalarStdDev(int nMinCount, float nMinValidValue, float nMaxValidValue, int nNumberOfRead, float nMinChangeRatio)
        {
            GenericTools.DebugMsg("GetScalarStdDev(nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + ", nMinChangeRatio=" + nMinChangeRatio.ToString() + "): Starting...");

            float nFirstStdDev = float.NaN;
            float nSecondStdDev = float.NaN;
            float nFirstAverage = float.NaN;
            float ReturnValue = float.NaN;

            DataTable StatsData = Connection.DataTable("avg(OverallValue) as AverageVal, stddev(OverallValue) as StdDevVal", "MeasReading", "PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and RowNum<=" + nMinCount.ToString() + " and (OverallValue between " + Convert.ToString(nMinValidValue) + " and " + Convert.ToString(nMaxValidValue) + ") order by ReadingId desc");
            if (StatsData.Rows.Count > 0)
            {
                nFirstStdDev = (float)Convert.ToDouble(StatsData.Rows[0]["AverageVal"]); // SQLtoFloat("stddev(OverallValue)", "MeasReading", "PointId=" + nPointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall") + " and RowNum<=" + nMinCount.ToString() + " and (OverallValue between " + nMinValidValue.ToString() + " and " + nMaxValidValue.ToString() + ") order by ReadingId desc");
                if (nNumberOfRead <= 1)
                    ReturnValue = nFirstStdDev;
                else
                {
                    nFirstAverage = (float)Convert.ToDouble(StatsData.Rows[0]["StdDevVal"]); // SQLtoFloat("avg(OverallValue)", "MeasReading", "PointId=" + nPointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall") + " and RowNum<=" + nMinCount.ToString() + " and (OverallValue between " + Convert.ToString(nMinValidValue) + " and " + Convert.ToString(nMaxValidValue) + ") order by ReadingId desc");
                    //if ((Math.Abs(nMinValidValue / (nFirstAverage - (2 * nFirstStdDev))) >= nMinChangeRatio) | (Math.Abs(nMaxValidValue / (nFirstAverage + (2 * nFirstStdDev))) >= nMinChangeRatio)) return nFirstStdDev;
                    nSecondStdDev = ScalarStdDev(nMinCount, (nFirstAverage - (2 * nFirstStdDev)), (nFirstAverage + (2 * nFirstStdDev)), nNumberOfRead - 1, nMinChangeRatio);
                    if (float.IsNaN(nSecondStdDev) | float.IsInfinity(nSecondStdDev) | (nSecondStdDev <= float.MinValue))
                        ReturnValue = nFirstStdDev;
                    else
                        ReturnValue = nSecondStdDev;
                }
            }

            GenericTools.DebugMsg("GetScalarStdDev(nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + ", nMinChangeRatio=" + nMinChangeRatio.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }



        public float[,] BandStats(float StartFreq, float EndFreq, uint nMinCount, float nMinValidValue, float nMaxValidValue, uint FreqType, uint nNumberOfRead)
        {
            GenericTools.DebugMsg("GetBandStats(StartFreq=" + StartFreq.ToString() + " , EndFreq=" + EndFreq.ToString() + ", nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", FreqType=" + FreqType.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + "): Starting...");

            int OverallCount = 0;
            int PeakCount = 0;
            float[] TmpValue = new float[2];
            float OverallValue = 0;
            float PeakValue = 0;
            float TotalOverallAverage = 0;
            float TotalOverallStdDev = 0;
            float TotalPeakAverage = 0;
            float TotalPeakStdDev = 0;
            float[] OverallValues = new float[nMinCount];
            float[] PeakValues = new float[nMinCount];
            float[,] FirstReturnValue = new float[2, 2];
            float[,] SecondReturnValue = new float[2, 2];
            float[,] ReturnValue = new float[2, 2];

            Measurement TmpMeasurement;
            //AnPoint Point = new AnPoint(ref TreeElem);

            uint nMeasId = Connection.SQLtoUInt("max(MeasId)", "MeasReading", "PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString());

            while (((OverallCount < nMinCount) | (PeakCount < nMinCount)) & (nMeasId > 0))
            {
                TmpMeasurement = new Measurement(Connection, nMeasId);
                TmpValue = TmpMeasurement.ReadingFFT.BandData(StartFreq, EndFreq, FreqType);
                OverallValue = TmpValue[0];
                PeakValue = TmpValue[1];
                if (OverallCount < nMinCount)
                {
                    if ((OverallValue >= (double)nMinValidValue) & (OverallValue <= (double)nMaxValidValue))
                    {
                        TotalOverallAverage += OverallValue;
                        OverallValues[OverallCount] = OverallValue;
                        OverallCount++;
                    }
                }
                if (PeakCount < nMinCount)
                {
                    if ((PeakValue >= (double)nMinValidValue) & (PeakValue <= (double)nMaxValidValue))
                    {
                        TotalPeakAverage += PeakValue;
                        PeakValues[PeakCount] = PeakValue;
                        PeakCount++;
                    }
                }
                nMeasId = Connection.SQLtoUInt("max(MeasId)", "MeasReading", "MeasId<" + nMeasId.ToString() + " and PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString());
            }

            double OverallAverage = (TotalOverallAverage / OverallCount);
            for (int i = 0; i < OverallCount; i++)
            {
                TotalOverallStdDev += (float)Math.Pow((OverallValues[i] - OverallAverage), 2);
            }
            FirstReturnValue[0, 0] = (float)OverallAverage; //Average
            FirstReturnValue[0, 1] = (float)Math.Sqrt(TotalOverallStdDev / OverallCount); //StdValue

            double PeakAverage = (TotalPeakAverage / PeakCount);
            for (int i = 0; i < PeakCount; i++)
            {
                TotalPeakStdDev += (float)Math.Pow((PeakValues[i] - PeakAverage), 2);
            }
            FirstReturnValue[1, 0] = (float)PeakAverage; //Average
            FirstReturnValue[1, 1] = (float)Math.Sqrt(TotalPeakStdDev / PeakCount); //StdValue

            if (nNumberOfRead <= 1)
            {
                ReturnValue = FirstReturnValue;
            }
            else
            {
                SecondReturnValue = BandStats(StartFreq, EndFreq, nMinCount, nMinValidValue, (FirstReturnValue[0, 0] + 2 * FirstReturnValue[0, 1]), FreqType, (nNumberOfRead - 1));

                if (float.IsNaN(SecondReturnValue[0, 0]) | float.IsInfinity(SecondReturnValue[0, 0]) | (SecondReturnValue[0, 0] <= float.MinValue))
                {
                    ReturnValue = FirstReturnValue;
                }
                else
                {
                    ReturnValue = SecondReturnValue;
                }
            }

            GenericTools.DebugMsg("GetBandStats(StartFreq=" + StartFreq.ToString() + " , EndFreq=" + EndFreq.ToString() + ", nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", FreqType=" + FreqType.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }

        public float[, ,] HarmonicBandStats(
                                                float StartFreq,
                                                float EndFreq,
                                                uint NumOfHarmonics,
                                                uint nMinCount,
                                                float nMinValidValue,
                                                float nMaxValidValue,
                                                uint FreqType,
                                                uint nNumberOfRead
                                            )
        {
            GenericTools.DebugMsg("GetHarmonicBandStats(StartFreq=" + StartFreq.ToString() + " , EndFreq=" + EndFreq.ToString() + ", nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", FreqType=" + FreqType.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + "): Starting...");

            int OverallCount0 = 0;
            int PeakCount0 = 0;
            int OverallCount1 = 0;
            int PeakCount1 = 0;
            int OverallCount2 = 0;
            int PeakCount2 = 0;
            int OverallCount3 = 0;
            int PeakCount3 = 0;
            int OverallCount4 = 0;
            int PeakCount4 = 0;
            int OverallCount5 = 0;
            int PeakCount5 = 0;
            float[,] TmpValue = new float[6, 2];
            float OverallValue0 = 0;
            float PeakValue0 = 0;
            float OverallValue1 = 0;
            float PeakValue1 = 0;
            float OverallValue2 = 0;
            float PeakValue2 = 0;
            float OverallValue3 = 0;
            float PeakValue3 = 0;
            float OverallValue4 = 0;
            float PeakValue4 = 0;
            float OverallValue5 = 0;
            float PeakValue5 = 0;
            float TotalOverallAverage0 = 0;
            float TotalOverallStdDev0 = 0;
            float TotalPeakAverage0 = 0;
            float TotalPeakStdDev0 = 0;
            float TotalOverallAverage1 = 0;
            float TotalOverallStdDev1 = 0;
            float TotalPeakAverage1 = 0;
            float TotalPeakStdDev1 = 0;
            float TotalOverallAverage2 = 0;
            float TotalOverallStdDev2 = 0;
            float TotalPeakAverage2 = 0;
            float TotalPeakStdDev2 = 0;
            float TotalOverallAverage3 = 0;
            float TotalOverallStdDev3 = 0;
            float TotalPeakAverage3 = 0;
            float TotalPeakStdDev3 = 0;
            float TotalOverallAverage4 = 0;
            float TotalOverallStdDev4 = 0;
            float TotalPeakAverage4 = 0;
            float TotalPeakStdDev4 = 0;
            float TotalOverallAverage5 = 0;
            float TotalOverallStdDev5 = 0;
            float TotalPeakAverage5 = 0;
            float TotalPeakStdDev5 = 0;
            float[] OverallValues0 = new float[nMinCount];
            float[] PeakValues0 = new float[nMinCount];
            float[] OverallValues1 = new float[nMinCount];
            float[] PeakValues1 = new float[nMinCount];
            float[] OverallValues2 = new float[nMinCount];
            float[] PeakValues2 = new float[nMinCount];
            float[] OverallValues3 = new float[nMinCount];
            float[] PeakValues3 = new float[nMinCount];
            float[] OverallValues4 = new float[nMinCount];
            float[] PeakValues4 = new float[nMinCount];
            float[] OverallValues5 = new float[nMinCount];
            float[] PeakValues5 = new float[nMinCount];
            float[, ,] FirstReturnValue = new float[6, 2, 2];
            float[, ,] SecondReturnValue = new float[6, 2, 2];
            float[, ,] ReturnValue = new float[6, 2, 2];

            Measurement TmpMeasurement;
            //AnPoint Point = new AnPoint(ref TreeElem);

            uint nMeasId = Connection.SQLtoUInt("max(MeasId)", "MeasReading", "PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString());

            while
                (
                    (
                        ((OverallCount0 < nMinCount) | (PeakCount0 < nMinCount)) |
                        ((OverallCount1 < nMinCount) | (PeakCount1 < nMinCount)) |
                        ((OverallCount2 < nMinCount) | (PeakCount2 < nMinCount)) |
                        ((OverallCount3 < nMinCount) | (PeakCount3 < nMinCount)) |
                        ((OverallCount4 < nMinCount) | (PeakCount4 < nMinCount)) |
                        ((OverallCount5 < nMinCount) | (PeakCount5 < nMinCount))
                     ) & (nMeasId > 0)
                )
            {
                TmpMeasurement = new Measurement(Connection, nMeasId);
                TmpValue = TmpMeasurement.ReadingFFT.HarmonicBandData(StartFreq, EndFreq, NumOfHarmonics, FreqType);
                OverallValue0 = TmpValue[0, 0];
                PeakValue0 = TmpValue[0, 1];
                OverallValue1 = TmpValue[1, 0];
                PeakValue1 = TmpValue[1, 1];
                OverallValue2 = TmpValue[2, 0];
                PeakValue2 = TmpValue[2, 1];
                OverallValue3 = TmpValue[3, 0];
                PeakValue3 = TmpValue[3, 1];
                OverallValue4 = TmpValue[4, 0];
                PeakValue4 = TmpValue[4, 1];
                OverallValue5 = TmpValue[5, 0];
                PeakValue5 = TmpValue[5, 1];
                if (OverallCount0 < nMinCount)
                {
                    if ((OverallValue0 >= (double)nMinValidValue) & (OverallValue0 <= (double)nMaxValidValue))
                    {
                        TotalOverallAverage0 += OverallValue0;
                        OverallValues0[OverallCount0] = OverallValue0;
                        OverallCount0++;
                    }
                }
                if (PeakCount0 < nMinCount)
                {
                    if ((PeakValue0 >= (double)nMinValidValue) & (PeakValue0 <= (double)nMaxValidValue))
                    {
                        TotalPeakAverage0 += PeakValue0;
                        PeakValues0[PeakCount0] = PeakValue0;
                        PeakCount0++;
                    }
                }
                if ((NumOfHarmonics >= 1) & (OverallCount1 < nMinCount))
                {
                    if ((OverallValue1 >= (double)nMinValidValue) & (OverallValue1 <= (double)nMaxValidValue))
                    {
                        TotalOverallAverage1 += OverallValue1;
                        OverallValues1[OverallCount1] = OverallValue1;
                        OverallCount1++;
                    }
                }
                if ((NumOfHarmonics >= 1) & (PeakCount1 < nMinCount))
                {
                    if ((PeakValue1 >= (double)nMinValidValue) & (PeakValue1 <= (double)nMaxValidValue))
                    {
                        TotalPeakAverage1 += PeakValue1;
                        PeakValues1[PeakCount1] = PeakValue1;
                        PeakCount1++;
                    }
                }
                if ((NumOfHarmonics >= 2) & (OverallCount2 < nMinCount))
                {
                    if ((OverallValue2 >= (double)nMinValidValue) & (OverallValue2 <= (double)nMaxValidValue))
                    {
                        TotalOverallAverage2 += OverallValue2;
                        OverallValues2[OverallCount2] = OverallValue2;
                        OverallCount2++;
                    }
                }
                if ((NumOfHarmonics >= 2) & (PeakCount2 < nMinCount))
                {
                    if ((PeakValue2 >= (double)nMinValidValue) & (PeakValue2 <= (double)nMaxValidValue))
                    {
                        TotalPeakAverage2 += PeakValue2;
                        PeakValues2[PeakCount2] = PeakValue2;
                        PeakCount2++;
                    }
                }
                if ((NumOfHarmonics >= 3) & (OverallCount3 < nMinCount))
                {
                    if ((OverallValue3 >= (double)nMinValidValue) & (OverallValue3 <= (double)nMaxValidValue))
                    {
                        TotalOverallAverage3 += OverallValue3;
                        OverallValues3[OverallCount3] = OverallValue3;
                        OverallCount3++;
                    }
                }
                if ((NumOfHarmonics >= 3) & (PeakCount3 < nMinCount))
                {
                    if ((PeakValue3 >= (double)nMinValidValue) & (PeakValue3 <= (double)nMaxValidValue))
                    {
                        TotalPeakAverage3 += PeakValue3;
                        PeakValues3[PeakCount3] = PeakValue3;
                        PeakCount3++;
                    }
                }
                if ((NumOfHarmonics >= 4) & (OverallCount4 < nMinCount))
                {
                    if ((OverallValue4 >= (double)nMinValidValue) & (OverallValue4 <= (double)nMaxValidValue))
                    {
                        TotalOverallAverage4 += OverallValue4;
                        OverallValues4[OverallCount4] = OverallValue4;
                        OverallCount4++;
                    }
                }
                if ((NumOfHarmonics >= 4) & (PeakCount4 < nMinCount))
                {
                    if ((PeakValue4 >= (double)nMinValidValue) & (PeakValue4 <= (double)nMaxValidValue))
                    {
                        TotalPeakAverage4 += PeakValue4;
                        PeakValues4[PeakCount4] = PeakValue4;
                        PeakCount4++;
                    }
                }
                if ((NumOfHarmonics >= 5) & (OverallCount5 < nMinCount))
                {
                    if ((OverallValue5 >= (double)nMinValidValue) & (OverallValue5 <= (double)nMaxValidValue))
                    {
                        TotalOverallAverage5 += OverallValue5;
                        OverallValues5[OverallCount5] = OverallValue5;
                        OverallCount5++;
                    }
                }
                if ((NumOfHarmonics >= 5) & (PeakCount5 < nMinCount))
                {
                    if ((PeakValue5 >= (double)nMinValidValue) & (PeakValue5 <= (double)nMaxValidValue))
                    {
                        TotalPeakAverage5 += PeakValue5;
                        PeakValues5[PeakCount5] = PeakValue5;
                        PeakCount5++;
                    }
                }
                nMeasId = Connection.SQLtoUInt("max(MeasId)", "MeasReading", "MeasId<" + nMeasId.ToString() + " and PointId=" + PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString());
            }

            double OverallAverage0 = (TotalOverallAverage0 / OverallCount0);
            for (int i = 0; i < OverallCount0; i++)
            {
                TotalOverallStdDev0 += (float)Math.Pow((OverallValues0[i] - OverallAverage0), 2);
            }
            FirstReturnValue[0, 0, 0] = (float)OverallAverage0; //Average
            FirstReturnValue[0, 0, 1] = (float)Math.Sqrt(TotalOverallStdDev0 / OverallCount0); //StdValue

            double PeakAverage0 = (TotalPeakAverage0 / PeakCount0);
            for (int i = 0; i < PeakCount0; i++)
            {
                TotalPeakStdDev0 += (float)Math.Pow((PeakValues0[i] - PeakAverage0), 2);
            }
            FirstReturnValue[0, 1, 0] = (float)PeakAverage0; //Average
            FirstReturnValue[0, 1, 1] = (float)Math.Sqrt(TotalPeakStdDev0 / PeakCount0); //StdValue

            double OverallAverage1 = (TotalOverallAverage1 / OverallCount1);
            for (int i = 0; i < OverallCount1; i++)
            {
                TotalOverallStdDev1 += (float)Math.Pow((OverallValues1[i] - OverallAverage1), 2);
            }
            FirstReturnValue[1, 0, 0] = (float)OverallAverage1; //Average
            FirstReturnValue[1, 0, 1] = (float)Math.Sqrt(TotalOverallStdDev1 / OverallCount1); //StdValue

            double PeakAverage1 = (TotalPeakAverage1 / PeakCount1);
            for (int i = 0; i < PeakCount1; i++)
            {
                TotalPeakStdDev1 += (float)Math.Pow((PeakValues1[i] - PeakAverage1), 2);
            }
            FirstReturnValue[1, 1, 0] = (float)PeakAverage1; //Average
            FirstReturnValue[1, 1, 1] = (float)Math.Sqrt(TotalPeakStdDev1 / PeakCount0); //StdValue

            double OverallAverage2 = (TotalOverallAverage2 / OverallCount2);
            for (int i = 0; i < OverallCount2; i++)
            {
                TotalOverallStdDev2 += (float)Math.Pow((OverallValues2[i] - OverallAverage2), 2);
            }
            FirstReturnValue[2, 0, 0] = (float)OverallAverage2; //Average
            FirstReturnValue[2, 0, 1] = (float)Math.Sqrt(TotalOverallStdDev2 / OverallCount2); //StdValue

            double PeakAverage2 = (TotalPeakAverage2 / PeakCount2);
            for (int i = 0; i < PeakCount2; i++)
            {
                TotalPeakStdDev2 += (float)Math.Pow((PeakValues2[i] - PeakAverage2), 2);
            }
            FirstReturnValue[2, 1, 0] = (float)PeakAverage2; //Average
            FirstReturnValue[2, 1, 1] = (float)Math.Sqrt(TotalPeakStdDev2 / PeakCount2); //StdValue

            double OverallAverage3 = (TotalOverallAverage3 / OverallCount3);
            for (int i = 0; i < OverallCount3; i++)
            {
                TotalOverallStdDev3 += (float)Math.Pow((OverallValues3[i] - OverallAverage3), 2);
            }
            FirstReturnValue[3, 0, 0] = (float)OverallAverage3; //Average
            FirstReturnValue[3, 0, 1] = (float)Math.Sqrt(TotalOverallStdDev3 / OverallCount3); //StdValue

            double PeakAverage3 = (TotalPeakAverage3 / PeakCount3);
            for (int i = 0; i < PeakCount3; i++)
            {
                TotalPeakStdDev3 += (float)Math.Pow((PeakValues3[i] - PeakAverage3), 2);
            }
            FirstReturnValue[3, 1, 0] = (float)PeakAverage3; //Average
            FirstReturnValue[3, 1, 1] = (float)Math.Sqrt(TotalPeakStdDev3 / PeakCount3); //StdValue

            double OverallAverage4 = (TotalOverallAverage4 / OverallCount4);
            for (int i = 0; i < OverallCount4; i++)
            {
                TotalOverallStdDev4 += (float)Math.Pow((OverallValues4[i] - OverallAverage4), 2);
            }
            FirstReturnValue[4, 0, 0] = (float)OverallAverage4; //Average
            FirstReturnValue[4, 0, 1] = (float)Math.Sqrt(TotalOverallStdDev4 / OverallCount4); //StdValue

            double PeakAverage4 = (TotalPeakAverage4 / PeakCount4);
            for (int i = 0; i < PeakCount4; i++)
            {
                TotalPeakStdDev4 += (float)Math.Pow((PeakValues4[i] - PeakAverage4), 2);
            }
            FirstReturnValue[4, 1, 0] = (float)PeakAverage4; //Average
            FirstReturnValue[4, 1, 1] = (float)Math.Sqrt(TotalPeakStdDev4 / PeakCount4); //StdValue

            double OverallAverage5 = (TotalOverallAverage5 / OverallCount5);
            for (int i = 0; i < OverallCount5; i++)
            {
                TotalOverallStdDev5 += (float)Math.Pow((OverallValues5[i] - OverallAverage5), 2);
            }
            FirstReturnValue[5, 0, 0] = (float)OverallAverage5; //Average
            FirstReturnValue[5, 0, 1] = (float)Math.Sqrt(TotalOverallStdDev5 / OverallCount5); //StdValue

            double PeakAverage5 = (TotalPeakAverage5 / PeakCount5);
            for (int i = 0; i < PeakCount5; i++)
            {
                TotalPeakStdDev5 += (float)Math.Pow((PeakValues5[i] - PeakAverage5), 2);
            }
            FirstReturnValue[5, 1, 0] = (float)PeakAverage5; //Average
            FirstReturnValue[5, 1, 1] = (float)Math.Sqrt(TotalPeakStdDev5 / PeakCount5); //StdValue

            if (nNumberOfRead <= 1)
                ReturnValue = FirstReturnValue;
            else
            {
                SecondReturnValue = HarmonicBandStats(StartFreq, EndFreq, NumOfHarmonics, nMinCount, nMinValidValue, (FirstReturnValue[0, 0, 0] + 2 * FirstReturnValue[0, 0, 1]), FreqType, (nNumberOfRead - 1));

                if (float.IsNaN(SecondReturnValue[0, 0, 0]) | float.IsInfinity(SecondReturnValue[0, 0, 0]) | (SecondReturnValue[0, 0, 0] <= float.MinValue))
                    ReturnValue = FirstReturnValue;
                else
                    ReturnValue = SecondReturnValue;
            }

            GenericTools.DebugMsg("GetHarmonicBandStats(StartFreq=" + StartFreq.ToString() + " , EndFreq=" + EndFreq.ToString() + ", nMinCount=" + nMinCount.ToString() + ", nMinValidValue=" + nMinValidValue.ToString() + ", nMaxValidValue=" + nMaxValidValue.ToString() + ", FreqType=" + FreqType.ToString() + ", nNumberOfRead=" + nNumberOfRead.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }
    }

    public class AnPoint
    {
        /*
        DataTable Points = new DataTable();

        private static Int32 SKFCM_ASPF_SCHEDULE_Type;
        private static Int32 SKFCM_ASPF_SCHEDULE_Take_Data;
        private static Int32 SKFCM_ASPF_SCHEDULE_Keep_Data;
        private static Int32 SKFCM_ASPF_SCHEDULE_Move_SArchive;
        private static Int32 SKFCM_ASPF_SCHEDULE_Move_LArchive;
        private static Int32 SKFCM_ASPF_SCHEDULE_Keep_SArchive;
        private static Int32 SKFCM_ASPF_SCHEDULE_Keep_LArchive;
        private static Int32 SKFCM_ASPF_SCHEDULE_Keep_Alarm;
        private static Int32 SKFCM_ASPF_SCHEDULE_Take_Data_Unit;
        private static Int32 SKFCM_ASPF_SCHEDULE_Keep_Data_Unit;
        private static Int32 SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit;
        private static Int32 SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit;
        private static Int32 SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit;
        private static Int32 SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit;
        private static Int32 SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit;

        private static Int32 NumberOfPoints;
        */
        //private static bool IsLoaded = false;

        public AnalystConnection Connection;
        //public TreeElem TreeElem;

        public AnPoint() { }
        public AnPoint(AnalystConnection AnTree)
        {
            //NumberOfPoints++;

            //AnConnection
            Connection = AnTree;
            /*
            OraConnection = AnTree.OraConnection;
            SQLConnection = AnTree.SQLConnection;
            MSAConnection = AnTree.MSAConnection;
            _SetConnected(AnTree.IsConnected); //IsConnected = AnTree.IsConnected;
            Owner = AnTree.Owner;
            ConnectionString = AnTree.ConnectionString;
            DBName = AnTree.DBName;
            InitialCatalog = AnTree.InitialCatalog;
            User = AnTree.User;
            Password = AnTree.Password;
            DBType = AnTree.DBType;
            if (AnTree.IsConnected && !IsConnected) Connect();
            */
        }
        public AnPoint(ref TreeElem AnTree)
        {
            // NumberOfPoints++;

            //AnConnection
            Connection = AnTree.Connection;
            /*
            OraConnection = AnTree.OraConnection;
            SQLConnection = AnTree.SQLConnection;
            MSAConnection = AnTree.MSAConnection;
            _SetConnected(AnTree.IsConnected);//IsConnected = AnTree.IsConnected;
            Owner = AnTree.Owner;
            ConnectionString = AnTree.ConnectionString;
            DBName = AnTree.DBName;
            InitialCatalog = AnTree.InitialCatalog;
            User = AnTree.User;
            Password = AnTree.Password;
            DBType = AnTree.DBType;
            if (AnTree.IsConnected && !IsConnected) Connect();
            */

            //AnTreeElem
            /*
            _TreeElem = AnTree;
            TreeElem.TreeElemId = AnTree.TreeElemId;
            TreeElem.HierarchyId = AnTree.HierarchyId;
            TreeElem.ContainerType = AnTree.ContainerType;
            TreeElem.HierarchyType = AnTree.HierarchyType;
            TreeElem.ParentId = AnTree.ParentId;
            TreeElem.ParentRefId = AnTree.ParentRefId;
            TreeElem.Name = AnTree.Name;
            TreeElem.Description = AnTree.Description;
             */
        }
        /*
        public AnPoint(int AnDBType, string AnDBName, string AnDBInitialCatalog, string AnDBUser, string AnDBPass, Int32 AnPointId, bool AnArchive, bool AnDebug)
        {
            GenericTools.Debug = AnDebug;
            MessageBox.Show("Teste");
            AnConnection AnalystConnection = new AnConnection(AnDBType, AnDBName, AnDBInitialCatalog, AnDBUser, AnDBPass);
            if (!AnalystConnection.IsConnected) AnalystConnection.Connect();
            if (AnalystConnection.IsConnected)
            {
                NumberOfPoints++;
                ReloadPoint(AnPointId);
                if (AnArchive) DoArchive(AnPointId);
            }
        }
        */

        /*
        public void ReloadPoint()
        {
            ReloadPoint(TreeElem.TreeElemId);
        }
        public void ReloadPoint(uint nTreeElemId)
        {
            GenericTools.DebugMsg("ReloadPoint(" + nTreeElemId.ToString() + "): Starting...");
            DateTime StartTime = System.DateTime.Now;
            try
            {
                DataTable TreeElemTable = AnConnection.DataTable("ContainerType, HierarchyId, HierarchyType, ParentId, ParentRefId, Name, Description","TreeElem","TreeElemId=" + nTreeElemId.ToString());
                if (TreeElemTable.Rows.Count > 0)
                {
                    Int32 ContainerTypeTmp = Convert.ToInt32(TreeElemTable.Rows[0]["ContainerType"]); //SQLtoInt("ContainerType", "TreeElem", "TreeElemId=" + nTreeElemId.ToString());
                    if (ContainerTypeTmp > 0)
                    {
                        ContainerType = ContainerTypeTmp;
                        if (ContainerType == 4)
                        {
                            TreeElemId = nTreeElemId;
                            HierarchyId = Convert.ToInt32(TreeElemTable.Rows[0]["HierarchyId"]); //, "TreeElem", "TreeElemId=" + TreeElemId.ToString());
                            HierarchyType = Convert.ToInt32(TreeElemTable.Rows[0]["HierarchyType"]); //, "TreeElem", "TreeElemId=" + TreeElemId.ToString());
                            ParentId = Convert.ToInt32(TreeElemTable.Rows[0]["ParentId"]); //, "TreeElem", "TreeElemId=" + TreeElemId.ToString());
                            ParentRefId = Convert.ToInt32(TreeElemTable.Rows[0]["ParentRefId"]); //, "TreeElem", "TreeElemId=" + TreeElemId.ToString());
                            Name = TreeElemTable.Rows[0]["Name"].ToString(); //, "TreeElem", "TreeElemId=" + TreeElemId.ToString());
                            Description = TreeElemTable.Rows[0]["Description"].ToString(); //, "TreeElem", "TreeElemId=" + TreeElemId.ToString());

                            SKFCM_ASPF_SCHEDULE_Type = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Type").ToString());
                            SKFCM_ASPF_SCHEDULE_Take_Data = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data").ToString());
                            SKFCM_ASPF_SCHEDULE_Keep_Data = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data").ToString());
                            SKFCM_ASPF_SCHEDULE_Move_SArchive = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive").ToString());
                            SKFCM_ASPF_SCHEDULE_Move_LArchive = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive").ToString());
                            SKFCM_ASPF_SCHEDULE_Keep_SArchive = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive").ToString());
                            SKFCM_ASPF_SCHEDULE_Keep_LArchive = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive").ToString());
                            SKFCM_ASPF_SCHEDULE_Keep_Alarm = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm").ToString());
                            SKFCM_ASPF_SCHEDULE_Take_Data_Unit = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Take_Data_Unit").ToString());
                            SKFCM_ASPF_SCHEDULE_Keep_Data_Unit = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Data_Unit").ToString());
                            SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_SArchive_Unit").ToString());
                            SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Move_LArchive_Unit").ToString());
                            SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_SArchive_Unit").ToString());
                            SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_LArchive_Unit").ToString());
                            SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit = SQLtoInt("ValueString", "Point", "ElementId=" + TreeElemId.ToString() + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_SCHEDULE_Keep_Alarm_Unit").ToString());

                            IsLoaded = true;
                        }
                    }
                    else
                    {
                        GenericTools.DebugMsg("ReloadPoint(" + nTreeElemId + "): Point not found - " + (System.DateTime.Now - StartTime).ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.GetError(ex.Message);
                GenericTools.DebugMsg("ReloadPoint(" + nTreeElemId.ToString() + ") error:" + ex.Message);
            }
            GenericTools.DebugMsg("ReloadPoint(" + nTreeElemId + "): " + (System.DateTime.Now - StartTime).ToString() + " secs");
        }
        */

        /*
        public Int32 GetNumberOfPoints()
        {
            return NumberOfPoints;
        }
        */

    }
   
    public class Measurement
    {
        //private bool _IsLoaded = false;
        public bool IsLoaded { get { return (MeasurementeRow != null); } }

        private Point _Point;
        public Point Point { get { return _Point; } }
        public uint PointId { get { return _Point.PointId; } }

        public TreeElem TreeElem { get { return (IsLoaded ? _Point.TreeElem : null); } }
        public AnalystConnection Connection { get { return (IsLoaded ? _Point.TreeElem.Connection : null); } }

        private uint _MeasId;
        public uint MeasId { get { return (IsLoaded ? _MeasId : 0); } }

        public DateTime DataDTG { get { return (IsLoaded ? GenericTools.StrToDateTime(_MeasurmentRow["DataDTG"].ToString()) : new DateTime()); } }
        public string OperName { get { return (IsLoaded ? _MeasurmentRow["OperName"].ToString() : string.Empty); } }
        public string SerialNo { get { return (IsLoaded ? _MeasurmentRow["SerialNo"].ToString() : string.Empty); } }
        public MeasurementStatus Status { get { return (IsLoaded ? (MeasurementStatus)_MeasurmentRow["Status"] : MeasurementStatus.None); } }

        private bool _FFT = false;
        public bool FFT
        {
            get
            {
                if (IsLoaded) _FFT = (Connection.SQLtoUInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString()) > 0);
                if (_FFT) _ReadingFFT = new MeasReadingFFT(this);
                return _FFT;
            }
        }
        private MeasReadingFFT _ReadingFFT;
        public MeasReadingFFT ReadingFFT { get { return (FFT ? _ReadingFFT : null); } }
        public bool Overall { get { return OverallReading.Overall; } }
        private MeasReadingOverall _OverallReading = null;
        public MeasReadingOverall OverallReading
        {
            get
            {
                if (IsLoaded & (_OverallReading == null)) _OverallReading = new MeasReadingOverall(Connection, MeasId);
                return _OverallReading;
            }
        }
        public bool Inspection { get { return InspectionReading.Inspection; } }
        private MeasReadingInspection _InspectionReading;
        public MeasReadingInspection InspectionReading
        {
            get
            {
                if (IsLoaded & (_InspectionReading == null)) _InspectionReading = new MeasReadingInspection(Connection, MeasId);
                return _InspectionReading;
            }
        }

        private bool _MCD = false;
        public bool MCD
        {
            get
            {

                if (IsLoaded) _MCD = (Connection.SQLtoUInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_MCD").ToString()) > 0);

                return _MCD;
            }
        }
        private MeasReadingMCD _MCDReading;
        public MeasReadingMCD MCDReading
        {
            get
            {

                if (IsLoaded & (_MCDReading == null)) _MCDReading = new MeasReadingMCD(Connection, MeasId);
                return _MCDReading;
            }
        }
        private Measurement _NextMeas = null;
        public Measurement NextMeas
        {
            get
            {
                if (_NextMeas == null)
                {
                    string MinDataDtg = Connection.SQLtoString("min(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and DataDtg>'" + GenericTools.DateTime(DataDTG) + "'");
                    if (!string.IsNullOrEmpty(MinDataDtg))
                    {
                        uint NextMeasId = Connection.SQLtoUInt("min(MeasId)", "Measurement", "PointId=" + PointId.ToString() + " and DataDtg='" + MinDataDtg + "'");
                        if (NextMeasId > 0) _NextMeas = new Measurement(Connection, NextMeasId);
                    }
                }
                return _NextMeas;
            }
        }
        private Measurement _PrevMeas = null;
        public Measurement PrevMeas
        {
            get
            {
                if (_PrevMeas == null)
                {
                    string MaxDataDtg = Connection.SQLtoString("max(DataDtg)", "Measurement", "PointId=" + PointId.ToString() + " and DataDtg<'" + GenericTools.DateTime(DataDTG) + "'");
                    if (!string.IsNullOrEmpty(MaxDataDtg))
                    {
                        uint PrevMeasId = Connection.SQLtoUInt("min(MeasId)", "Measurement", "PointId=" + PointId.ToString() + " and DataDtg='" + MaxDataDtg + "'");
                        if (PrevMeasId > 0) _PrevMeas = new Measurement(Connection, PrevMeasId);
                    }
                }
                return _PrevMeas;
            }
        }
        private DataRow _MeasurmentRow = null;
        public DataRow MeasurementeRow { get { return _MeasurmentRow; } }
        public Measurement() { }
        public Measurement(AnalystConnection AnalystConnection, uint MeasId)
        {
            _Point = new Point(AnalystConnection, AnalystConnection.SQLtoUInt("PointId", "Measurement", "MeasId=" + MeasId.ToString()));
            if (_Point.TreeElem.IsLoaded)
            {
                _MeasId = MeasId;
                _Load();
            }
        }
        public Measurement(Point Point)
        {
            this._Point = Point;
            if (_Point.TreeElem.IsLoaded)
            {
                _MeasId = _Point.LastMeas.MeasId;
                _Load();
            }
        }
        private bool _Load()
        {
            _MeasurmentRow = null;

            if (_Point.TreeElem.IsLoaded && _MeasId > 0)
            {
                DataTable MeasurementRows = _Point.Connection.DataTable("*", "Measurement", "MeasId=" + _MeasId.ToString());
                if (MeasurementRows.Rows.Count > 0)
                    _MeasurmentRow = MeasurementRows.Rows[0];
            }

            return IsLoaded;
        }

        public DataTable Notes = new DataTable();

        public string GetNotesString()
        {
            return GetNotesString(";");
        }
        public string GetNotesString(string Delimiter)
        {
            GenericTools.DebugMsg("GetNotesString(" + Delimiter + "): Starting...");
            string ReturnValue = string.Empty;

            try
            {
                for (Int32 i = 0; i < Notes.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + Delimiter;
                    ReturnValue = ReturnValue + Notes.Rows[i][0].ToString();
                }
                ReturnValue = ReturnValue.Replace(Environment.NewLine, Delimiter).Trim();

                while (ReturnValue.IndexOf("  ") > 0)
                    ReturnValue = ReturnValue.Replace("  ", " ");

                ReturnValue = ReturnValue.Trim();
            }
            catch (Exception ex)
            {
                GenericTools.GetError("GetNotesString(" + Delimiter + ") = " + ReturnValue + ": " + ex.Message);
            }

            GenericTools.DebugMsg("GetNotesString(" + Delimiter + "): " + ReturnValue);
            return ReturnValue;
        }

    }

    public class AnMeasurement : Measurement
    {
        public uint _MeasId;
        public new uint MeasId = 0;
        //        public int ReadingType = 0;
        //public DateTime DataDTG = DateTime.MinValue;
        public string Operator = string.Empty;

        public float OverallValue = float.NaN;
        public string OverallFullScaleUnits = string.Empty;

        /*
        public bool FFT = false;
        public Int32 FFTReadingId;
        public string FFTReadingHeader;
        public int FFTLines;
        public float FFTResolution;
        public float FFTMaxFrequency;
        public int FFTAverages;
        public float FFTSpeed;
        public float FFTFactor;
        public float FFTOverallRatio = 1.2237f;
        public float[] FFTReadingData;
        */

        public bool Inspection = false;
        public Int32 InspectionReadingId;
        public int InspectionResult;
        public string[] InspectionOptions;
        public int InspectionAlarmFlag;


        public bool MCD = false;
        public MeasMCDOld MeasMCD;
        /*
        public bool MCD = false;
        public Int32 MCDReadingId;
        public float MCDEnvelope;
        public float MCDVelocity;
        public float MCDTemperature;
        */

        //public DataTable Notes = new DataTable();

        public bool Loaded = false;


        private static Int32 NumberOfMeasurements;

        /*
        public AnMeasurement(AnalystConnection AnalystConnection, uint MeasId)
        {
            ReloadMeas(MeasId);
        }
        */

        /*
        public AnMeasurement(ref AnPoint AnPt)
        {
            //AnConnection
            Connection = AnPt.Connection;
            
            OraConnection = AnPt.OraConnection;
            SQLConnection = AnPt.SQLConnection;
            MSAConnection = AnPt.MSAConnection;
            _SetConnected(AnPt.IsConnected);
            Owner = AnPt.Owner;
            ConnectionString = AnPt.ConnectionString;
            DBName = AnPt.DBName;
            InitialCatalog = AnPt.InitialCatalog;
            User = AnPt.User;
            Password = AnPt.Password;
            DBType = AnPt.DBType;
            if (AnPt.IsConnected) Connect();

            //AnTreeElem
            TreeElemId = AnPt.TreeElemId;
            HierarchyId = AnPt.HierarchyId;
            ContainerType = AnPt.ContainerType;
            HierarchyType = AnPt.HierarchyType;
            ParentId = AnPt.ParentId;
            ParentRefId = AnPt.ParentRefId;
            Name = AnPt.Name;
            Description = AnPt.Description;
        }
        */

        /*
        public void ReloadMeas()
        {
            ReloadMeas(MeasId);
        }
        public void ReloadMeas(uint MeasId)
        {
            GenericTools.DebugMsg("ReloadMeas(" + MeasId.ToString() + "): Starting...");

            try
            {
                DataTable MeasIdOverallDTG;
                string tDataDTG = string.Empty;

                if (Connection.SQLtoUInt("count(*)", "MeasReading", "MeasId=" + MeasId.ToString()) > 0)
                {
                    _MeasId = MeasId;
                    TreeElemId = SQLtoInt("PointId", "Measurement", "MeasId=" + MeasId.ToString());
                    Operator = SQLtoString("OperName", "Measurement", "MeasId=" + MeasId.ToString());
                    tDataDTG = SQLtoString("DataDTG", "Measurement", "MeasId=" + MeasId.ToString());
                    DataDTG = GenericTools.StrToDateTime(tDataDTG);//GenericTools.StrToDateTime(MeasIdOverallDTG.Rows[0]["DataDtg"].ToString()); //

                    OverallFullScaleUnits = string.Empty;

                    MeasIdOverallDTG = RecordSet("select OverallValue, DataDtg, OperName from " + Owner + "MeasDTSRead where ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and MeasId=" + MeasId.ToString());
                    if (MeasIdOverallDTG.Rows.Count > 0)
                    {
                        OverallValue = (float)Convert.ToDouble(MeasIdOverallDTG.Rows[0]["OverallValue"]); //SQLtoFloat("OverallValue", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString());
                        OverallFullScaleUnits = SQLtoString("ValueString", "Point", "ElementId=" + TreeElemId + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit").ToString());
                    }
                    MeasIdOverallDTG = RecordSet("select ReadingId, ReadingHeader, ReadingData from " + Owner + "MeasReading where ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString() + " and MeasId=" + MeasId.ToString());

                    FFT = (MeasIdOverallDTG.Rows.Count > 0);//(SQLtoInt("count(*)", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString()) > 0);
                    if (FFT)
                    {
                        FFTReadingId = Convert.ToInt32(MeasIdOverallDTG.Rows[0]["ReadingId"]); //SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString());
                        FFTReadingHeader = BLOBtoString(MeasIdOverallDTG.Rows[0]["ReadingHeader"]); //select ReadingHeader from " + Owner + "MeasReading where ReadingId=" + FFTReadingId.ToString());
                        FFTFactor = GenericTools.Hex8ToValue(GenericTools.ReverseHex(FFTReadingHeader.Substring(48, 16)));
                        FFTReadingData = GetBLOBtoArray(MeasIdOverallDTG.Rows[0]["ReadingData"], FFTFactor);
                        FFTSpeed = GenericTools.Hex8ToValue(GenericTools.ReverseHex(FFTReadingHeader.Substring(32, 16)));
                        FFTMaxFrequency = GenericTools.Hex8ToValue(GenericTools.ReverseHex(FFTReadingHeader.Substring(80, 16)));
                        FFTLines = FFTReadingData.Length; //SQLtoInt("(((dbms_lob.getLength(ReadingData))/2)-1)", "MeasReading", "ReadingId=" + FFTReadingId.ToString());
                        // FFTAverages = 
                        FFTResolution = (float)(FFTMaxFrequency / FFTLines);
                    }

                    InspectionReadingId = SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection").ToString());
                    Inspection = (InspectionReadingId > 0);
                    if (Inspection)
                    {
                        InspectionResult = SQLtoInt("ExDWordVal1", "MeasReading", "ReadingId=" + InspectionReadingId.ToString());
                        InspectionOptions = GetInspection();
                        InspectionAlarmFlag = GetInspectionResultAlarmFlag();
                    }
                    else
                    {
                        InspectionResult = int.MinValue;
                        InspectionOptions = null;
                    }

                    AnConnection Analyst = new AnConnection();
                    Analyst.OraConnection = OraConnection;
                    Analyst.SQLConnection = SQLConnection;
                    Analyst.MSAConnection = MSAConnection;
                    Analyst._SetConnected(IsConnected); //Analyst.IsConnected = IsConnected;
                    Analyst.Owner = Owner;
                    Analyst.ConnectionString = ConnectionString;
                    Analyst.DBName = DBName;
                    Analyst.InitialCatalog = InitialCatalog;
                    Analyst.User = User;
                    Analyst.Password = Password;
                    Analyst.DBType = DBType;
                    if (IsConnected && !Analyst.IsConnected) Analyst.Connect();

                    //MCDReadingId = SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_MCD").ToString());
                    MCD = (SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_MCD").ToString()) > 0); //(MCDReadingId > 0);
                    if (MCD)
                    {
                        MeasMCD = new AnMeasMCD(Analyst, MeasId);
                    }
                    /*
                        MCDEnvelope = SQLtoFloat("ExDoubleVal3", "MeasReading", "ReadingId=" + MCDReadingId.ToString());
                        MCDVelocity = SQLtoFloat("ExDoubleVal2", "MeasReading", "ReadingId=" + MCDReadingId.ToString());
                        MCDTemperature = SQLtoFloat("ExDoubleVal1", "MeasReading", "ReadingId=" + MCDReadingId.ToString());
                        OverallFullScaleUnits = SQLtoString("ValueString", "Point", "ElementId=" + TreeElemId + " and FieldId=" + Registration.RegistrationId(Connection, "SKFCM_ASPF_Full_Scale_Unit").ToString());
                    }
                    else
                    {
                        MCDEnvelope = float.NaN;
                        MCDVelocity = float.NaN;
                        MCDTemperature = float.NaN;
                    }
                    */
        /*
                    Notes = RecordSet("Text", "Notes", "OwnerId=" + TreeElemId.ToString() + " and DataDTG between '" + GenericTools.DateTime(DataDTG.AddDays(-1)) + "' and '" + GenericTools.DateTime(DataDTG.AddDays(+1)) + "'");

                    Loaded = true;
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("ReloadMeas(" + nMeasId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("ReloadMeas(" + nMeasId.ToString() + "): Finished!");
        }
        */

        /*
        public static AnMeasurement AddMeasurement(AnMeasurement AnMeas, uint nMeasId)
        {
            AnMeas.MeasId = nMeasId;

            if (AnMeas.TreeElemId < 1)
            {
                AnMeas.TreeElemId = AnMeas.SQLtoInt("PointId", "Measurement", "MeasId=" + nMeasId.ToString());
                AnMeas.ReloadMeas();
            }

            AnMeas.Resset(nMeasId);

            for (Int32 i = 0; i < NumberOfMeasurements; i++)
            {
                if (MeasReadingElements[i].MeasId == nMeasId) return MeasReadingElements[i];
            }

            NumberOfMeasurements++;

            if (NumberOfMeasurements > MeasReadingElements.GetLength(0)) Array.Resize<AnMeasurement>(ref MeasReadingElements, NumberOfMeasurements);

            MeasReadingElements[NumberOfMeasurements - 1] = AnMeas;

            return AnMeas;
        }
        */

        /*
        public Int32 GetNumberOfMeasurements()
        {
            return NumberOfMeasurements;
        }
        */

        /*
        public void Resset(Int32 nMeasId)
        {
            MeasId = nMeasId;
            OverallValue = GetOverall(nMeasId);
            _ReadFFT(nMeasId);
        }
        */

        //public string GetNotesString()
        //{
        //    return GetNotesString(";");
        //}
        //public string GetNotesString(string Delimiter)
        //{
        //    GenericTools.DebugMsg("GetNotesString(" + Delimiter + "): Starting...");
        //    string ReturnValue = string.Empty;

        //    try
        //    {
        //        for (Int32 i = 0; i < Notes.Rows.Count; i++)
        //        {
        //            if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + Delimiter;
        //            ReturnValue = ReturnValue + Notes.Rows[i][0].ToString();
        //        }
        //        ReturnValue = ReturnValue.Replace(Environment.NewLine, Delimiter).Trim();

        //        while (ReturnValue.IndexOf("  ") > 0)
        //            ReturnValue = ReturnValue.Replace("  ", " ");

        //        ReturnValue = ReturnValue.Trim();
        //    }
        //    catch (Exception ex)
        //    {
        //        GenericTools.GetError("GetNotesString(" + Delimiter + ") = " + ReturnValue + ": " + ex.Message);
        //    }

        //    GenericTools.DebugMsg("GetNotesString(" + Delimiter + "): " + ReturnValue);
        //    return ReturnValue;
        //}
    }
    public class MeasMCDOld
    {
        private Int32 iTreeElemId = 0;
        private Int32 iMeasId = 0;

        public DateTime TimeStamp;

        public MeasReadingOverall EnvOverallValue;
        public MeasReadingOverall VelOverallValue;
        public MeasReadingOverall TempOverallValue;

        public bool TempMeasured;

        public MeasReadingFFT EnvFFT;
        public MeasReadingFFT VelFFT;

        public AlarmMCD _AlarmLevel;
        public MCDMeasAlarm MCDMeasAlarm;

        private AnConnection Analyst;


        /*
        static public void AnMeasMCD(AnConnection AnalystConnection)
        {
            AnMeasMCD(AnalystConnection, 0);
        }
        */
        /**public void AnMeasMCD(AnConnection AnalystConnection, Int32 MeasId)
        {
            GenericTools.DebugMsg("AnMeasMCD(" + MeasId.ToString() + "): Starting...");
            Analyst = AnalystConnection;
            if (MeasId>0) Load(MeasId);
            GenericTools.DebugMsg("AnMeasMCD(" + MeasId.ToString() + "): Finished");
        }
        **/

        private void Initialize()
        {
            GenericTools.DebugMsg("Initialize(): Starting...");

            iTreeElemId = 0;
            iMeasId = 0;
            TempMeasured = false;

            GenericTools.DebugMsg("Initialize(): Initializing DateTime...");
            TimeStamp = new DateTime();

            GenericTools.DebugMsg("Initialize(): Initializing AnMeasOverall...");
            EnvOverallValue = new MeasReadingOverall();
            VelOverallValue = new MeasReadingOverall();
            TempOverallValue = new MeasReadingOverall();

            GenericTools.DebugMsg("Initialize(): Initializing AnMeasFFT...");
            EnvFFT = new MeasReadingFFT();
            VelFFT = new MeasReadingFFT();

            GenericTools.DebugMsg("Initialize(): Initializing AnAlarmMCD...");
            _AlarmLevel = new AlarmMCD();
            GenericTools.DebugMsg("Initialize(): Initializing AnMCDMeasAlarm...");
            MCDMeasAlarm = new MCDMeasAlarm();

            GenericTools.DebugMsg("Initialize(): Finished!");
        }


        /**public bool Load(Int32 MeasId)
        {
            GenericTools.DebugMsg("Load("+ MeasId.ToString() +"): Starting...");

            try
            {
                Initialize();

                if (Analyst.IsConnected)
                {
                    DataTable Measurement = Analyst.DataTable("PointId, Status, Include, OperName, SerialNo, DataDTG, MeasurementType", "Measurement", "MeasId=" + MeasId.ToString());

                    if (Measurement.Rows.Count > 0)
                    {
                        iTreeElemId = Int32.Parse(Measurement.Rows[0]["PointId"].ToString());
                        TimeStamp = GenericTools.StrToDateTime(Measurement.Rows[0]["DataDTG"].ToString());

                        DataTable MeasReading = Analyst.DataTable("ReadingId, ExDWordVal1, ExDoubleVal1, ExDoubleVal2, ExDoubleVal3", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Analyst.GetRegistrationId("SKFCM_ASMD_MCD").ToString());

                        if (MeasReading.Rows.Count > 0)
                        {
                            GenericTools.DebugMsg("Load(" + MeasId.ToString() + "): " + MeasReading.Rows[0]["ExDoubleVal3"].ToString());
                            EnvOverallValue.OverallValue = GenericTools.StrToFloat(MeasReading.Rows[0]["ExDoubleVal3"].ToString());
                            GenericTools.DebugMsg("Load(" + MeasId.ToString() + "): EnvOverallValue.OverallValue=" + EnvOverallValue.OverallValue.ToString());
                            VelOverallValue.OverallValue = GenericTools.StrToFloat(MeasReading.Rows[0]["ExDoubleVal2"].ToString());
                            GenericTools.DebugMsg("Load(" + MeasId.ToString() + "): VelOverallValue.OverallValue=" + VelOverallValue.OverallValue.ToString());
                            TempMeasured = (MeasReading.Rows[0]["ExDWordVal1"].ToString() == "7");
                            GenericTools.DebugMsg("Load(" + MeasId.ToString() + "): TempMeasured=" + TempMeasured.ToString());

                            if (TempMeasured)
                                TempOverallValue.OverallValue = GenericTools.StrToFloat(MeasReading.Rows[0]["ExDoubleVal1"].ToString());
                            GenericTools.DebugMsg("Load(" + MeasId.ToString() + "): TempOverallValue.OverallValue=" + TempOverallValue.OverallValue.ToString());

                            iMeasId = MeasId;

                            LoadOverallUnits();
                            LoadAlarmLevel();
                            CalcAlarm();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Load(" + MeasId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("Load(" + MeasId.ToString() + "): " + (iMeasId > 0).ToString());

            return (iMeasId > 0);
        }
        **/

        public AlarmMCD LoadAlarmLevel()
        {
            _AlarmLevel = LoadAlarmLevel(iTreeElemId);
            return _AlarmLevel;
        }
        /// <summary>
        /// ADICIONADO POR MATEUS - VISITA VALE - 24/06/2014 - UTILIZADO NO ODR REPORT CONSOLE
        /// </summary>
        /// <param name="PointId"></param>
        /// <returns></returns>
        /// 
        public AlarmMCD LoadAlarmLevel(AnalystConnection Analyst, uint PointId)
        {
            GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + "): Starting...");
            AlarmMCD ReturnValue = new AlarmMCD();

            try
            {
                if (Analyst.IsConnected)
                {
                    int AlarmId = Analyst.SQLtoInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString() + " and Type=" + Registration.RegistrationId(Analyst, "SKFCM_ASAT_MCD"));
                    DataTable tMCDAlarm = Analyst.DataTable("ALARMSET, ENVACCELENABLEDANGERHI, ENVACCELENABLEALERTHI, ENVACCELDANGERHI, ENVACCELALERTHI, TEMPERATUREENABLEDANGERHI, TEMPERATUREENABLEALERTHI, TEMPERATUREDANGERHI, TEMPERATUREALERTHI, VELOCITYENABLEDANGERHI, VELOCITYENABLEALERTHI, VelocityDangerHi, VelocityAlertHi", "MCDAlarm", ((AlarmId > 0) ? ("MCDAlarmId=" + AlarmId.ToString()) : ("ElementId=" + PointId.ToString())));

                    if (tMCDAlarm.Rows.Count > 0)
                    {
                        ReturnValue.Envelope.AlarmMethod = AlarmOverallMethod.Level;
                        ReturnValue.Envelope.EnableDangerHi = (tMCDAlarm.Rows[0]["EnvAccelEnableDangerHi"].ToString() == "1");
                        ReturnValue.Envelope.EnableAlertHi = (tMCDAlarm.Rows[0]["EnvAccelEnableAlertHi"].ToString() == "1");
                        ReturnValue.Envelope.EnableAlertLo = false;
                        ReturnValue.Envelope.EnableDangerLo = false;
                        ReturnValue.Envelope.DangerHi = float.Parse(tMCDAlarm.Rows[0]["EnvAccelDangerHi"].ToString());
                        ReturnValue.Envelope.AlertHi = float.Parse(tMCDAlarm.Rows[0]["EnvAccelAlertHi"].ToString());
                        ReturnValue.Envelope.AlertLo = 0;
                        ReturnValue.Envelope.DangerLo = 0;

                        ReturnValue.Velocity.AlarmMethod = AlarmOverallMethod.Level;
                        ReturnValue.Velocity.EnableDangerHi = (tMCDAlarm.Rows[0]["VelocityEnableDangerHi"].ToString() == "1");
                        ReturnValue.Velocity.EnableAlertHi = (tMCDAlarm.Rows[0]["VelocityEnableAlertHi"].ToString() == "1");
                        ReturnValue.Velocity.EnableAlertLo = false;
                        ReturnValue.Velocity.EnableDangerLo = false;
                        ReturnValue.Velocity.DangerHi = float.Parse(tMCDAlarm.Rows[0]["VelocityDangerHi"].ToString());
                        ReturnValue.Velocity.AlertHi = float.Parse(tMCDAlarm.Rows[0]["VelocityAlertHi"].ToString());
                        ReturnValue.Velocity.AlertLo = 0;
                        ReturnValue.Velocity.DangerLo = 0;

                        ReturnValue.Temperature.AlarmMethod = AlarmOverallMethod.Level;
                        ReturnValue.Temperature.EnableDangerHi = (tMCDAlarm.Rows[0]["TemperatureEnableDangerHi"].ToString() == "1");
                        ReturnValue.Temperature.EnableAlertHi = (tMCDAlarm.Rows[0]["TemperatureEnableAlertHi"].ToString() == "1");
                        ReturnValue.Temperature.EnableAlertLo = false;
                        ReturnValue.Temperature.EnableDangerLo = false;
                        ReturnValue.Temperature.DangerHi = float.Parse(tMCDAlarm.Rows[0]["TemperatureDangerHi"].ToString());
                        ReturnValue.Temperature.AlertHi = float.Parse(tMCDAlarm.Rows[0]["TemperatureAlertHi"].ToString());
                        ReturnValue.Temperature.AlertLo = 0;
                        ReturnValue.Temperature.DangerLo = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }
        public AlarmMCD LoadAlarmLevel(Int32 PointId)
        {
            GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + "): Starting...");
            AlarmMCD ReturnValue = new AlarmMCD();

            try
            {
                if (Analyst.IsConnected)
                {
                    int AlarmId = Analyst.SQLtoInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString() + " and Type=" + Analyst.GetRegistrationId("SKFCM_ASAT_MCD"));
                    DataTable tMCDAlarm = Analyst.DataTable("ALARMSET, ENVACCELENABLEDANGERHI, ENVACCELENABLEALERTHI, ENVACCELDANGERHI, ENVACCELALERTHI, TEMPERATUREENABLEDANGERHI, TEMPERATUREENABLEALERTHI, TEMPERATUREDANGERHI, TEMPERATUREALERTHI, VELOCITYENABLEDANGERHI, VELOCITYENABLEALERTHI", "MCDAlarm", ((AlarmId > 0) ? ("MCDAlarmId=" + AlarmId.ToString()) : ("ElementId=" + PointId.ToString())));

                    if (tMCDAlarm.Rows.Count > 0)
                    {
                        ReturnValue.Envelope.AlarmMethod = AlarmOverallMethod.Level;
                        ReturnValue.Envelope.EnableDangerHi = (tMCDAlarm.Rows[0]["EnvAccelEnableDangerHi"].ToString() == "1");
                        ReturnValue.Envelope.EnableAlertHi = (tMCDAlarm.Rows[0]["EnvAccelEnableAlertHi"].ToString() == "1");
                        ReturnValue.Envelope.EnableAlertLo = false;
                        ReturnValue.Envelope.EnableDangerLo = false;
                        ReturnValue.Envelope.DangerHi = float.Parse(tMCDAlarm.Rows[0]["EnvAccelDangerHi"].ToString());
                        ReturnValue.Envelope.AlertHi = float.Parse(tMCDAlarm.Rows[0]["EnvAccelAlertHi"].ToString());
                        ReturnValue.Envelope.AlertLo = 0;
                        ReturnValue.Envelope.DangerLo = 0;

                        ReturnValue.Velocity.AlarmMethod = AlarmOverallMethod.Level;
                        ReturnValue.Velocity.EnableDangerHi = (tMCDAlarm.Rows[0]["VelocityEnableDangerHi"].ToString() == "1");
                        ReturnValue.Velocity.EnableAlertHi = (tMCDAlarm.Rows[0]["VelocityEnableAlertHi"].ToString() == "1");
                        ReturnValue.Velocity.EnableAlertLo = false;
                        ReturnValue.Velocity.EnableDangerLo = false;
                        ReturnValue.Velocity.DangerHi = float.Parse(tMCDAlarm.Rows[0]["VelocityDangerHi"].ToString());
                        ReturnValue.Velocity.AlertHi = float.Parse(tMCDAlarm.Rows[0]["VelocityAlertHi"].ToString());
                        ReturnValue.Velocity.AlertLo = 0;
                        ReturnValue.Velocity.DangerLo = 0;

                        ReturnValue.Temperature.AlarmMethod = AlarmOverallMethod.Level;
                        ReturnValue.Temperature.EnableDangerHi = (tMCDAlarm.Rows[0]["TemperatureEnableDangerHi"].ToString() == "1");
                        ReturnValue.Temperature.EnableAlertHi = (tMCDAlarm.Rows[0]["TemperatureEnableAlertHi"].ToString() == "1");
                        ReturnValue.Temperature.EnableAlertLo = false;
                        ReturnValue.Temperature.EnableDangerLo = false;
                        ReturnValue.Temperature.DangerHi = float.Parse(tMCDAlarm.Rows[0]["TemperatureDangerHi"].ToString());
                        ReturnValue.Temperature.AlertHi = float.Parse(tMCDAlarm.Rows[0]["TemperatureAlertHi"].ToString());
                        ReturnValue.Temperature.AlertLo = 0;
                        ReturnValue.Temperature.DangerLo = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("LoadAlarmLevel(" + PointId.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }


        /**        public void LoadOverallUnits()
        {
            GenericTools.DebugMsg("LoadOverallUnits(): Starting...");
            if (iTreeElemId > 0 && Analyst.IsConnected)
            {
                MeasReadingOverall MCDOverall = new MeasReadingOverall();
                string MCDUnits = MCDOverall.LoadFullScaleUnit(Analyst, iTreeElemId);

                if (MCDUnits.IndexOf(',') > 0)
                {
                    EnvOverallValue.Point.FullScaleUnit = MCDUnits.Split(',')[0].Trim();
                    GenericTools.DebugMsg("LoadOverallUnits(): EnvOverallValue.FullScaleUnit = " + EnvOverallValue.FullScaleUnit);
                    VelOverallValue.Point.FullScaleUnit = MCDUnits.Split(',')[1].Trim();
                    GenericTools.DebugMsg("LoadOverallUnits(): VelOverallValue.FullScaleUnit = " + VelOverallValue.FullScaleUnit);
                    TempOverallValue.Point.FullScaleUnit = MCDUnits.Split(',')[2].Trim();
                    GenericTools.DebugMsg("TempOverallValue(): EnvOverallValue.FullScaleUnit = " + TempOverallValue.FullScaleUnit);
                }
            }
            GenericTools.DebugMsg("LoadOverallUnits(): Starting...");
        }
        **/

        public MCDMeasAlarm CalcAlarm()
        {
            MCDMeasAlarm = CalcAlarm(iMeasId);
            return MCDMeasAlarm;
        }
        /// <summary>
        /// ADICIONADO POR MATEUS - VISITA VALE 24/06/2014
        /// NECESSÁRIO PARA O ODR REPORT CONSOLE. USANDO CONEXÃO NOVA.
        /// </summary>
        /// <param name="Connect"></param>
        /// <param name="MeasId"></param>
        /// <returns></returns>
        //public MCDMeasAlarm CalcAlarm(AnalystConnection Connect, uint MeasId)
        public MCDMeasAlarm CalcAlarm(Measurement Measure)
        {

            GenericTools.DebugMsg("CalcAlarm(" + Measure.MeasId.ToString() + "): Starting...");

            MCDMeasAlarm ReturnValue = new MCDMeasAlarm();

            MeasMCDOld AlMCD = new MeasMCDOld();
            _AlarmLevel = AlMCD.LoadAlarmLevel(Measure.Connection, Measure.PointId);

            try
            {
                ReturnValue.Envelope.AlarmType = Convert.ToInt32(Registration.RegistrationId(Measure.Connection, "SKFCM_ASAT_MCD"));
                ReturnValue.Velocity.AlarmType = ReturnValue.Envelope.AlarmType;
                ReturnValue.Temperature.AlarmType = ReturnValue.Envelope.AlarmType;

                double Envelope_Overall = Convert.ToDouble(Measure.MCDReading.MeasReadingRow["EXDOUBLEVAL3"]);
                double Velocidade_Overall = Convert.ToDouble(Measure.MCDReading.MeasReadingRow["EXDOUBLEVAL2"]);
                double Temperatura_Overall = Convert.ToDouble(Measure.MCDReading.MeasReadingRow["EXDOUBLEVAL1"]);

                ReturnValue.Envelope.AlarmLevel = AlarmLevel.Good;
                if ((_AlarmLevel.Envelope.EnableAlertHi) && (_AlarmLevel.Envelope.AlertHi <= Envelope_Overall)) ReturnValue.Envelope.AlarmLevel = AlarmLevel.Alert;
                if ((_AlarmLevel.Envelope.EnableDangerHi) && (_AlarmLevel.Envelope.DangerHi <= Envelope_Overall)) ReturnValue.Envelope.AlarmLevel = AlarmLevel.Danger;

                ReturnValue.Velocity.AlarmLevel = AlarmLevel.Good;
                if ((_AlarmLevel.Velocity.EnableAlertHi) && (_AlarmLevel.Velocity.AlertHi <= Velocidade_Overall)) ReturnValue.Velocity.AlarmLevel = AlarmLevel.Alert;
                if ((_AlarmLevel.Velocity.EnableDangerHi) && (_AlarmLevel.Velocity.DangerHi <= Velocidade_Overall)) ReturnValue.Velocity.AlarmLevel = AlarmLevel.Danger;

                //if (TempMeasured)
                //{
                ReturnValue.Temperature.AlarmLevel = AlarmLevel.Good;
                if ((_AlarmLevel.Temperature.EnableAlertHi) && (_AlarmLevel.Temperature.AlertHi <= Temperatura_Overall)) ReturnValue.Temperature.AlarmLevel = AlarmLevel.Alert;
                if ((_AlarmLevel.Temperature.EnableDangerHi) && (_AlarmLevel.Temperature.DangerHi <= Temperatura_Overall)) ReturnValue.Temperature.AlarmLevel = AlarmLevel.Danger;
                //}
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("CalcAlarm(" + Measure.MeasId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("CalcAlarm(" + Measure.MeasId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public MCDMeasAlarm CalcAlarm(Int32 MeasId)
        {
            GenericTools.DebugMsg("CalcAlarm(" + MeasId.ToString() + "): Starting...");

            MCDMeasAlarm ReturnValue = new MCDMeasAlarm();

            try
            {
                if (Analyst.IsConnected)
                {
                    ReturnValue.Envelope.AlarmType = Convert.ToInt32(Analyst.GetRegistrationId("SKFCM_ASAT_MCD"));
                    ReturnValue.Velocity.AlarmType = ReturnValue.Envelope.AlarmType;
                    ReturnValue.Temperature.AlarmType = ReturnValue.Envelope.AlarmType;


                    ReturnValue.Envelope.AlarmLevel = AlarmLevel.Good;
                    if ((_AlarmLevel.Envelope.EnableAlertHi) && (_AlarmLevel.Envelope.AlertHi <= EnvOverallValue.OverallValue)) ReturnValue.Envelope.AlarmLevel = AlarmLevel.Alert;
                    if ((_AlarmLevel.Envelope.EnableDangerHi) && (_AlarmLevel.Envelope.DangerHi <= EnvOverallValue.OverallValue)) ReturnValue.Envelope.AlarmLevel = AlarmLevel.Danger;

                    ReturnValue.Velocity.AlarmLevel = AlarmLevel.Good;
                    if ((_AlarmLevel.Velocity.EnableAlertHi) && (_AlarmLevel.Velocity.AlertHi <= VelOverallValue.OverallValue)) ReturnValue.Velocity.AlarmLevel = AlarmLevel.Alert;
                    if ((_AlarmLevel.Velocity.EnableDangerHi) && (_AlarmLevel.Velocity.DangerHi <= VelOverallValue.OverallValue)) ReturnValue.Velocity.AlarmLevel = AlarmLevel.Danger;

                    if (TempMeasured)
                    {
                        ReturnValue.Temperature.AlarmLevel = AlarmLevel.Good;
                        if ((_AlarmLevel.Temperature.EnableAlertHi) && (_AlarmLevel.Velocity.AlertHi <= TempOverallValue.OverallValue)) ReturnValue.Temperature.AlarmLevel = AlarmLevel.Alert;
                        if ((_AlarmLevel.Temperature.EnableDangerHi) && (_AlarmLevel.Velocity.DangerHi <= TempOverallValue.OverallValue)) ReturnValue.Temperature.AlarmLevel = AlarmLevel.Danger;
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("CalcAlarm(" + MeasId.ToString() + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("CalcAlarm(" + MeasId.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
    }
    public class AlarmMCD
    {
        public AlarmOverall Envelope;
        public AlarmOverall Velocity;
        public AlarmOverall Temperature;

        public AlarmMCD()
        {
            Envelope = new AlarmOverall();
            Velocity = new AlarmOverall();
            Temperature = new AlarmOverall();
        }
    }
    public class AlarmOverall
    {
        public AlarmOverallMethod AlarmMethod;

        public bool EnableDangerHi;
        public bool EnableDangerLo;
        public bool EnableAlertHi;
        public bool EnableAlertLo;

        public float DangerHi;
        public float DangerLo;
        public float AlertHi;
        public float AlertLo;

    }
    public class MeasAlarm
    {
        public Int32 AlarmType = 0;
        public AlarmLevel AlarmLevel = AlarmLevel.Null;
        public Int32 ResultType = 0;
        public Int32 Channel = 0;

        public MeasAlarm()
        {
            AlarmType = 0;
            AlarmLevel = AlarmLevel.Null;
            ResultType = 0;
            Channel = 0;
        }
    }
    public class MCDMeasAlarm
    {
        public MeasAlarm Envelope = new MeasAlarm();
        public MeasAlarm Velocity = new MeasAlarm();
        public MeasAlarm Temperature = new MeasAlarm();

        public MCDMeasAlarm()
        {
            GenericTools.DebugMsg("AnMCDMeasAlarm(): Starting...");

            GenericTools.DebugMsg("AnMCDMeasAlarm(): Initializing Channel...");
            Envelope.Channel = 1;
            Velocity.Channel = 1;
            Temperature.Channel = 1;

            GenericTools.DebugMsg("AnMCDMeasAlarm(): Initializing ResultType...");
            Envelope.ResultType = 1;
            Velocity.ResultType = 2;
            Temperature.ResultType = 3;

            GenericTools.DebugMsg("AnMCDMeasAlarm(): Initializing AlarmLevel...");
            Envelope.AlarmLevel = AlarmLevel.None;
            Velocity.AlarmLevel = AlarmLevel.None;
            Temperature.AlarmLevel = AlarmLevel.None;
        }
    }
    public class MeasReading
    {
        //private bool _IsLoaded = false;
        public bool IsLoaded { get { return Measurement.IsLoaded; } }

        public DataRow MeasReadingRow;

        private Measurement _Measurement;
        public Measurement Measurement { get { return _Measurement; } }//  set { _Measurement = value; } }
        public AnalystConnection Connection { get { return (IsLoaded ? _Measurement.Connection : null); } }
        public uint MeasId { get { return (IsLoaded ? _Measurement.MeasId : new uint()); } }
        public uint PointId { get { return (IsLoaded ? _Measurement.PointId : new uint()); } }
        public DateTime DataDTG { get { return (IsLoaded ? _Measurement.DataDTG : new DateTime()); } }
        public MeasurementStatus Status { get { return (IsLoaded ? _Measurement.Status : MeasurementStatus.None); } }

        private uint _ReadingId;
        public uint ReadingId { get { return (IsLoaded ? _ReadingId : new uint()); } set { _ReadingId = value; } }

        private ReadingType _Type = ReadingType.None;
        public ReadingType Type { get { return (IsLoaded ? _Type : ReadingType.None); } }

        private uint _Channel = 0;
        public uint Channel { get { return (IsLoaded ? _Channel : new uint()); } }

        public MeasReading() { }
        public MeasReading(Measurement mMeasurement)
        {
            _Measurement = mMeasurement;
        }


    }
    public class MeasReadingBOV : MeasReading { }
    public class MeasReadingBaseline : MeasReading { }
    public class MeasReadingFFT
    {
        //private bool _IsLoaded = false;
        public bool IsLoaded { get { return MeasReading.IsLoaded; } }

        private MeasReading _MeasReading = null;
        public MeasReading MeasReading { get { return _MeasReading; } }

        public AnalystConnection Connection { get { return MeasReading.Connection; } }

        private string _ReadingHeaderText;
        public string ReadingHeaderText { get { return (MeasReading.IsLoaded ? _ReadingHeaderText : string.Empty); } }

        private float[] _ReadingHeaderBlob;
        public float[] ReadingHeaderBlob { get { return (MeasReading.IsLoaded ? _ReadingHeaderBlob : null); } }

        public uint _Lines;
        public uint Lines { get { return (MeasReading.IsLoaded ? _Lines : 0); } }
        public uint FFTLines { get { return (MeasReading.IsLoaded ? _Lines : 0); } }

        private float _Resolution;
        public float Resolution { get { return (MeasReading.IsLoaded ? _Resolution : float.NaN); } }

        private float _MaxFrequency;
        public float MaxFrequency { get { return (MeasReading.IsLoaded ? _MaxFrequency : float.NaN); } }

        private float _LowFrequency;
        public float LowFrequency { get { return (MeasReading.IsLoaded ? _LowFrequency : float.NaN); } }


        public uint _Averages;
        public uint Averages { get { return (MeasReading.IsLoaded ? _Averages : 0); } }

        private float _Speed;
        public float Speed { get { return (MeasReading.IsLoaded ? _Speed : float.NaN); } }

        private float _Factor;
        //public float Factor { get { return (_IsLoaded ? _Factor : float.NaN); } }

        private float _OverallRatio = 1.2237f;

        private float[] _ReadingData;
        public float[] ReadingData { get { return (MeasReading.IsLoaded ? _ReadingData : null); } }

        private bool _SyncAndNonsyncLoaded = false;
        private float[] _SyncAndNonsync = null;
        public float SyncOverall
        {
            get
            {
                try
                {
                    if (!_SyncAndNonsyncLoaded)
                        SyncAndNonsync();
                    if (_SyncAndNonsync != null)
                        return _SyncAndNonsync[0];
                }
                catch { }
                return float.NaN;
            }
        }
        public float NonsyncOverall
        {
            get
            {
                try
                {
                    if (!_SyncAndNonsyncLoaded)
                        SyncAndNonsync();
                    if (_SyncAndNonsync != null)
                        return _SyncAndNonsync[1];
                }
                catch { }
                return float.NaN;
            }
        }

        public MeasReadingFFT() { }
        public MeasReadingFFT(Measurement Measurement)
        {
            _MeasReading = new Analyst.MeasReading(Measurement);
            _Load();
        }
        public MeasReadingFFT(AnalystConnection AnalystConnection, uint MeasId)
        {
            _MeasReading = new Analyst.MeasReading(new Measurement(AnalystConnection, MeasId));
            _Load();
        }
        public MeasReadingFFT(Point Point)
        {
            //uint MeasId = Point.Connection.SQLtoUInt("max(MeasId)", "Measurement", "PointId=" + Point.PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Point.Connection, "SKFCM_ASMD_FFT").ToString());
            uint MeasId = Point.Connection.SQLtoUInt("max(DataDTG)", "Measurement", "PointId=" + Point.PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Point.Connection, "SKFCM_ASMD_FFT").ToString());
            _MeasReading = new Analyst.MeasReading(new Measurement(Point));
            _Load();
        }

        private bool _Load()
        {
            //_IsLoaded = false;

            if (MeasReading.Measurement.IsLoaded)
            {
                DataTable MeasReadingRows = MeasReading.Measurement.Connection.DataTable("ReadingId, ReadingHeader, ReadingData", "MeasReading", "MeasId=" + MeasReading.Measurement.MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(MeasReading.Measurement.Connection, "SKFCM_ASMD_FFT").ToString());
                if (MeasReadingRows.Rows.Count > 0)
                {
                    MeasReading.MeasReadingRow = MeasReadingRows.Rows[0];
                    MeasReading.ReadingId = Convert.ToUInt32(MeasReading.MeasReadingRow["ReadingId"]); //SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString());
                    _ReadingHeaderText = Connection.BLOBtoString(MeasReading.MeasReadingRow["ReadingHeader"]); //select ReadingHeader from " + Owner + "MeasReading where ReadingId=" + FFTReadingId.ToString());
                    _Factor = GenericTools.Hex8ToValue(GenericTools.ReverseHex(_ReadingHeaderText.Substring(48, 16)));
                    List<float> HeaderCreator = new List<float>();
                    for (int i = 0; i < (int)(_ReadingHeaderText.Length / 16); i++)
                    {
                        HeaderCreator.Add(GenericTools.Hex8ToValue(GenericTools.ReverseHex(_ReadingHeaderText.Substring(16 * i, 16))));
                    }
                    _ReadingHeaderBlob = HeaderCreator.ToArray<float>();// Connection.GetBLOBtoArray(MeasReading.MeasReadingRow["ReadingHeader"], 1);
                    _ReadingData = Connection.GetBLOBtoArray(MeasReading.MeasReadingRow["ReadingData"], _Factor);
                    _ReadingData = _ReadingData.Skip(1).ToArray();
                    _Speed = GenericTools.Hex8ToValue(GenericTools.ReverseHex(_ReadingHeaderText.Substring(32, 16)));
                    _MaxFrequency = GenericTools.Hex8ToValue(GenericTools.ReverseHex(_ReadingHeaderText.Substring(80, 16)));
                    //_MaxFrequency = GenericTools.Hex8ToValue(GenericTools.ReverseHex(_ReadingHeaderText.Substring(80, 16)));
                    _Lines = Convert.ToUInt32(_ReadingData.Length); // Math.Max(0, Convert.ToUInt32(_ReadingData.Length) - 1); //SQLtoInt("(((dbms_lob.getLength(ReadingData))/2)-1)", "MeasReading", "ReadingId=" + FFTReadingId.ToString());
                    _Averages = 0;
                    // FFTAverages = 
                    if (_Lines > 0)
                        _Resolution = (float)(_MaxFrequency / _Lines);
                    else
                        _Resolution = float.NaN;

                    _SyncAndNonsyncLoaded = false;
                    _SyncAndNonsync = null;


                    //_IsLoaded = true;
                }
            }

            return IsLoaded;
        }


        public float GetAmplitude(uint Line) { return (IsLoaded ? _ReadingData[Line] : float.NaN); }
        //public float GetAmplitude() { return _Factor; }

        public float BandPeak(uint FirstLine, uint LastLine) { return BandData(FirstLine, LastLine)[1]; }
        public float BandPeak(uint Line) { return BandData(Line)[1]; }
        public float BandPeak() { return BandData()[1]; }

        public float BandOverall(uint FirstLine, uint LastLine) { return BandData(FirstLine, LastLine)[0]; }
        public float BandOverall(uint Line) { return BandData(Line)[0]; }
        public float BandOverall() { return BandData()[0]; }

        public float[] BandData() { return BandData(1, _Lines); }
        public float[] BandData(uint Line) { return BandData(Line, Line); }
        public float[] BandData(uint FirstLine, uint LastLine)
        {
            GenericTools.DebugMsg("BandData(" + FirstLine.ToString() + ", " + LastLine.ToString() + "): Starting...");
            if (!IsLoaded) return null;

            float[] ReturnValue = new float[2];
            float ReturnOverall = 0;
            float ReturnPeak = 0;

            uint Line1 = Math.Min(FirstLine, LastLine);
            uint Line2 = Math.Max(FirstLine, LastLine);

            if (Line1 < 1) Line1 = 1;
            if (Line2 > FFTLines) Line2 = FFTLines;

            if (Line1 == Line2)
            {
                ReturnValue[0] = _ReadingData[Line1 - 1];
                ReturnValue[1] = ReturnValue[0];
            }
            else
            {
                for (uint i = Line1; i <= Line2; i++)
                {
                    ReturnOverall += (float)Math.Pow(_ReadingData[i - 1], 2);
                    ReturnPeak = Math.Max(ReturnPeak, _ReadingData[i - 1]);
                }
                ReturnValue[0] = (float)Math.Sqrt(ReturnOverall / _OverallRatio);
                ReturnValue[1] = ReturnPeak;
            }
            GenericTools.DebugMsg("BandData(" + FirstLine.ToString() + ", " + LastLine.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        public float[] BandData(float StartFreq, float EndFreq, uint FreqType)
        {
            GenericTools.DebugMsg("BandData(" + StartFreq.ToString() + ", " + EndFreq.ToString() + ", " + FreqType.ToString() + "): Starting...");
            float[] ReturnValue = new float[2];

            if ((_Resolution > 0) & (StartFreq >= 0) & (EndFreq >= StartFreq))
            {
                if (FreqType == (uint)SKF.RS.STB.Analyst.FreqType.Order)
                {
                    StartFreq = StartFreq * _Speed;
                    EndFreq = EndFreq * _Speed;
                }
                ReturnValue = BandData(Convert.ToUInt32(Math.Max(Math.Min(Math.Round((StartFreq / _Resolution), 0), UInt32.MaxValue), UInt32.MinValue)), Convert.ToUInt32(Math.Max(Math.Min(Math.Round((EndFreq / _Resolution), 0), UInt32.MaxValue), UInt32.MinValue)));
            }
            GenericTools.DebugMsg("BandData(" + StartFreq.ToString() + ", " + EndFreq.ToString() + ", " + FreqType.ToString() + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        public float[] BandData(float nFrequency, uint FreqType) { return BandData(nFrequency, nFrequency, FreqType); }
        //private float[] _BandData() { return BandData(1, _Lines); }

        public float[,] HarmonicBandData(float StartFreq, float EndFreq, uint NumOfHarmonics, uint FreqType)
        {
            NumOfHarmonics = (uint)Math.Max(1, NumOfHarmonics);
            float[,] ReturnValue = new float[6, 2];
            float[] BandValue = BandData(StartFreq, EndFreq, FreqType);
            ReturnValue[0, 0] = BandValue[0];
            ReturnValue[0, 1] = BandValue[1];

            for (uint i = 1; i <= NumOfHarmonics; i++)
            {
                BandValue = BandData(StartFreq * (i + 1), EndFreq * (i + 1), FreqType);
                ReturnValue[i, 0] = BandValue[0];
                ReturnValue[i, 1] = BandValue[1];
            }

            return ReturnValue;
        }
        public float[,] HarmonicBandData(float Freq, uint NumOfHarmonics, uint FreqType)
        {
            return HarmonicBandData(Freq, Freq, NumOfHarmonics, FreqType);
        }

        /// <summary>
        /// Returns Sync and Unsync overall levels of FFT
        /// </summary>
        /// <returns>
        /// SyncAndUnsync[0] = Sync
        /// SyncAndUnsync[1] = Unsync
        /// </returns>
        public float[] SyncAndNonsync()
        {
            _SyncAndNonsyncLoaded = false;
            _SyncAndNonsync = null;

            if (!IsLoaded)
                return _SyncAndNonsync;

            if ((Resolution > 0) && (Speed > Resolution))
            {
                try
                {
                    uint LinRPM = (uint)(Math.Round(Speed / Resolution));

                    if (LinRPM > 1)
                    {
                        float ReturnSync = 0;
                        float ReturnUnsync = 0;

                        for (uint i = 1; i <= Lines; i++)
                        {
                            if ((i % LinRPM) == 0)
                                ReturnSync += (float)Math.Pow(_ReadingData[i - 1], 2);
                            else
                                ReturnUnsync += (float)Math.Pow(_ReadingData[i - 1], 2);
                        }
                        float[] TempSyncAndUnsync = { (float)Math.Sqrt(ReturnSync / _OverallRatio), (float)Math.Sqrt(ReturnUnsync / _OverallRatio) };
                        _SyncAndNonsync = TempSyncAndUnsync; //[0] = (float)Math.Sqrt(ReturnSync / _OverallRatio);
                        //_SyncAndUnsync[1] = (float)Math.Sqrt(ReturnUnsync / _OverallRatio);
                        _SyncAndNonsyncLoaded = true;
                    }
                }
                catch (Exception ex)
                {
                    _SyncAndNonsyncLoaded = false;
                    _SyncAndNonsync = null;
                }
            }
            return _SyncAndNonsync;
        }



        //----------------------------------------------------------------------------
        // DESC:
        //    This function calculates the hamonic activity index of the supplied spectrum
        //    at the location iFund with bandwidth iBandwidth. iBandwidth is applied to the
        //    highest harmonic in the spectrum and will cause this function to compute
        //    the index for a number of frequencies ranging between (iFund * N) +/- iBandwidth
        //    (with N being the max number of harmonics). Only the highest value for the
        //		index computed within this band is returns as "The Index". For further
        //    documentation refer to US patent 6,792,360 
        //
        //		oFund will return the fundamental frequency for which the highest index was
        //		found. This may be different from iFund as the search through iBandwidth may
        //		find a slightly different value for (iFund * N) delivering the highest index.
        //		The calling function may ignore oFund or use it to present harmonic activity
        //		with more finer detail.
        //
        //		If iBandwidth is set to zero, the index is computed for iFund only, no
        //		search is conducted and iFund == oFund always.
        //
        // PARAMS:
        //    iSpecLineIface - The Spectrum plot line interface
        //		xFundamental - The fundamental for which to compute the index (in Hz). This
        //			function may return a slightly changed value if it found a more optimal
        //			value in the search band (bandwidth).
        //    iBandwidth - Percentage, applied to highest harmonic value and used 
        //			to specify a 'search band' around the possible highest harmonic.
        //    iUnitType - As defined by CM_HAL_UT_XXX in HlprFunc.h
        //		oHai - a double being the Harmonic Activity Index for iFund. This value is 
        //			always positive.
        //
        // RETURNS:
        //		S_OK if the HAL value could be computed without issue
        //		S_FAIL if the HAL value could not be computed within the given values
        //    E_FAIL if there was one or more errors with supplied or derived interfaces
        //----------------------------------------------------------------------------
        public bool Harmonic_Activity_Index
           (
           double xFundamental,
           double iBandwidth,
           int iUnitType,
           out double oHAL
           )
        {
            int HAL_SEARCH_SAMPLES = 100;
            int CM_HAL_UT_VELOCITY = 2;
            int CM_HAL_UT_DISPLACEMENT = 3;

            int i;

            double HighFreq = this.MaxFrequency;
            int NumLines = (int)this.Lines;
            double[] FFTValues = Array.ConvertAll(_ReadingData, x => (double)x);
            double Resolution = this.Resolution;
            int Harmonics;
            double Klow, Khigh;
            double[] Energies = new double[HAL_SEARCH_SAMPLES];
            double[] Ratios = new double[HAL_SEARCH_SAMPLES]; // [ HAL_SEARCH_SAMPLES ];
            double[] LowEnergyCounter = new double[HAL_SEARCH_SAMPLES];


            // Calculate spectral resolution. Notice that spectra are stored with 
            // 1 extra bin (the zero/DC line) that should be ignored for calculating
            // the spectral resolution.
            //
            Resolution = HighFreq / (NumLines - 1);


            //
            // HAL calculations are all based on acceleration spectral. Based on the
            // unit type convert a velocity or displacement spectrum back to acceleration.
            // Notice that this does not need to be perfectly done with proper detection
            // etc. The most important is to get a spectral correction for increasingly
            // higher frequencies (i.e., differentiation)
            //
            if (iUnitType == CM_HAL_UT_VELOCITY)
                for (i = 0; i < (int)NumLines; i++)
                    FFTValues[i] *= 3.1415 * i * Resolution;
            else if (iUnitType == CM_HAL_UT_DISPLACEMENT)
                for (i = 0; i < (int)NumLines; i++)
                    FFTValues[i] *= 3.1415 * 3.1415 * (i * Resolution * i * Resolution);


            //
            // Calculate the amount of energy for the entire spectrum. 
            //
            double BaseEnergy = 0;
            for (i = 0; i < (int)NumLines; i++)
                BaseEnergy += FFTValues[i];

            //
            // Calculate the maximum number of harmonics of xFundamental possible
            // within HighFreq
            //
            Harmonics = (int)(HighFreq / xFundamental);

            //
            // Compute the minimum and maximum frequency values around
            // the maximum harmonic ( Harmonics * xFundamental ). This is the search band
            //
            Klow = (Harmonics * xFundamental) * (1 - iBandwidth / 100);
            Khigh = (Harmonics * xFundamental) * (1 + iBandwidth / 100);

            //
            // Make sure that Khigh is not higher than the maximum frequency.
            //
            if (Khigh > HighFreq)
                Khigh = HighFreq;

            for (i = 0; i < HAL_SEARCH_SAMPLES; i++)
            {
                Energies[i] = 0;
                Ratios[i] = 0;
                LowEnergyCounter[i] = 0;
            }

            //
            // Calculate the index values around iFund using HAL_SEARC_SAMPLES within
            // the search band defined by Klow and Khigh
            //
            double Freq;         // Holds frequency value of harmonic
            int HIdx;         // Harmonic counter
            double HEnerg;       // Holds total harmonic energy
            double Energy;       // Holds partial harmonic energy
            double PrevEnergy;   // Holds partial harmonic energy from previous harmonic
            double Line;         // Spectral line position i.e. 'bin'
            int LineBase;     // Integer part of 'Line'
            double LineFrac;     // Fractional part of 'Line'

            for (i = 1; i < HAL_SEARCH_SAMPLES; i++)
            {
                Freq = (xFundamental * Harmonics) + (i - HAL_SEARCH_SAMPLES / 2) *
                   (Khigh - Klow) / HAL_SEARCH_SAMPLES;
                Freq /= Harmonics;

                HIdx = 1;	   // Harmonics start at 1 (== fundamental)
                HEnerg = 0;	   // Accumulates harmonic energy for this iteration
                PrevEnergy = 0;
                while (HIdx <= Harmonics)
                {
                    //
                    // Calculate _exactly_ where the spectral line (index) is. Note that
                    // this might not be an integer value.
                    //
                    Line = (Freq * HIdx) / Resolution;
                    LineBase = (int)Line;
                    LineFrac = Line - LineBase;

                    //
                    // Calculate energy contribution for this 'f'. Notice that this is done
                    // by a linear interpolation between the spectral lines if 'Line' is not
                    // integer
                    //
                    Energy = 0; ;
                    if (LineBase < (int)NumLines - 1)
                    {
                        Energy = FFTValues[LineBase];
                        Energy += (FFTValues[LineBase + 1] - FFTValues[LineBase]) * LineFrac;
                    }

                    //
                    // Accumulate the harmonic energy only if the peak amplitude is larger
                    // then the average base energy. This makes sure that harmonic series with
                    // each peak being high are being favored over series that only have a few
                    // harmonics but generally display low spectral values.  Notice that if
                    // the energy for this peak is not accumulated because its value was too low,
                    // the harmonic index 'HIdx' continues to be incremented!
                    //
                    //Energy *= pHarmonic_Peak_Factor( FFTValues, NumLines, LineBase );
                    if (Energy > BaseEnergy / NumLines)
                        HEnerg += Energy * Harmonic_Peak_Factor(FFTValues, NumLines, LineBase);
                    else
                        LowEnergyCounter[i]++;

                    if (Energy > 0 && PrevEnergy > 0)
                    {
                        if (Energy < BaseEnergy / NumLines)
                            Ratios[i] += 2;
                        else
                            Ratios[i] += PrevEnergy / Energy;
                    }

                    PrevEnergy = Energy;

                    HIdx++;		// Next harmonic
                }

                //
                // Calculate average ratio. If harmonics amplitudes are more or less on the same
                // horizontal line or slowly move up or down, the average ratio will approach one.
                // However, if e.g., every other harmonic amplitude is low, high, low, high etc.
                // "Ratio" is significantly higher (or lower) then 1. Such event would signify that the 
                // fundamental frequency isn't the true frequency. 
                //
                Ratios[i] /= Harmonics;

                //
                // We have now accumulated all harmonic energy related for 'Freq' in Energies[]. This now
                // needs to be averaged across the number of harmonics ( Energies[] / Harmonics).  This in turn
                // must be divided by the average base energy (BaseEnergy / NumLines). However
                // BaseEnergy itself must be correct first, removing the harmonic energy collected
                // in Energies[]. Finally, the result is stored in the Energies[] array.
                //
                if (HIdx - 1 == 0 || NumLines - HIdx == 0)
                    Energies[i] = 0;
                else
                    Energies[i] = (HEnerg / (HIdx - 1)) / ((BaseEnergy - HEnerg) / (NumLines - HIdx));

                //
                // No need to go through the rest of the calculations if the bandwidth was set to zero
                //
                if (iBandwidth == 0)
                    break;
            }

            //
            // The memory was allocated by another COM object so use the 
            // COM safe memory management.
            //
            //::CoTaskMemFree( FFTValues );

            //
            // Find the maximum value in Energies[] and calculate the frequency that belongs to this
            //
            HIdx = 0;
            Energy = 0;
            for (i = 1; i < HAL_SEARCH_SAMPLES; i++)
                if (Energies[i] > Energy)
                {
                    Energy = Energies[i];
                    HIdx = i;

                    //
                    // No need to go through the rest of the calculations if the bandwidth was set to zero
                    //
                    if (iBandwidth == 0)
                        break;
                }

            xFundamental = (xFundamental * Harmonics) + (HIdx - HAL_SEARCH_SAMPLES / 2) *
                  (Khigh - Klow) / HAL_SEARCH_SAMPLES;
            xFundamental /= Harmonics;


            //
            // Calculate the Harmonic index. In essence Energies[] already holds that index
            // however, we want to add a few conditions;
            // a) Different harmonic series could have identical (or close) harmonic energy
            //    we want to favor those that have more harmonics.
            // b) A harmonic energy that was captured with Ratios[] being close to 1 (one)
            //    should be favored over those that aren't. Ratios[] >> 1 may indicate that
            //    the fundamental was a false 1/2, 1/3, etc down. Ratios[] << 1 may indicate
            //    that the fundamental was actually the 2X, 3X etc of the real fundamental.
            //    Since the value in Ratios[] is not strongly discriminant we use the
            //    square instead of the straight value.
            // c) Finally to bring the overall index down to a more or less linear operating
            //    value we take the square root of the computation.
            //
            if (Ratios[HIdx] >= 1)
                oHAL = Math.Sqrt(Energies[HIdx] * Harmonics / Ratios[HIdx]);
            else if (Ratios[HIdx] < 1)
                oHAL = Math.Sqrt(Energies[HIdx] * Harmonics * Ratios[HIdx]);
            else
                oHAL = Math.Sqrt(Energies[HIdx] * Harmonics);

            oHAL *= (Harmonics - LowEnergyCounter[HIdx]) / Harmonics;

            return true;
        }



        //----------------------------------------------------------------------------
        // DESC:
        //    This function calculates a factor that is a measure for how well defined
        //    the spectral peak at iLineBase is. Spectral content whereby the amplitude
        //    at iLineBase is clearly the highest, and is surrounded by a more or less
        //    gradual downwards slope is favored. Spectral content that is on par with
        //    the amplitude found at iLineBase is dampened. Spectral content that shows
        //    strong excursions beyond the amplitude at iLineBase is strongly suppressed.
        //
        //    What we want to determine is how much a spectral peak looks like:
        //
        //             *
        //            *** 
        //           *****
        //          *******
        //             ^
        //    versus:
        //
        //          *    *
        //          *  * **
        //          *******
        //          *******
        //             ^
        // PARAMS:
        //    iFFTValues, an array of doubles representing the FFT
        //    iNumLines, a int representing the length of iFFTValues[]
        //    iLineBase, the spectral line index where our supposed peak is located.
        //    xDurationString - The formated string as output.
        //----------------------------------------------------------------------------
        double Harmonic_Peak_Factor
              (
              double[] iFFTValues,
              int iNumLines,
              int iLineBase
              )
        {
            int i, Window, Left, Right;
            double AccDifference = 1;
            double Ratio = 0;


            //
            // Calculate the number of lines we'll be investigating left and right of
            // iLineBase. This is a percentage of the total number of lines. More lines
            // allow a peak to be better defined with more spectral lines surrounding it.
            // For the case of 100 lines, we should have at least 1 line on either side
            // i.e., the constant factor to use is 100, expecting this may be assumed to
            // be a linearly scaling problem for higher number of line spectra.
            //
            if (iNumLines < 1600)
                Window = 2;
            else if (iNumLines < 12800)
                Window = 3;
            else
                Window = 4;

            if ((int)iLineBase < Window)
                Left = 0;
            else
                Left = iLineBase - Window;

            if ((int)iLineBase + Window > iNumLines - 1)
                Right = iNumLines - 1;
            else
                Right = iLineBase + Window;

            //
            // Left side
            //
            for (i = Left; i < (int)iLineBase && i < iNumLines; i++)
            {
                Ratio = iFFTValues[i + 1] / iFFTValues[i];

                if (Ratio > 2)
                    AccDifference += 1.5;
                else if (Ratio < 0.5)
                    AccDifference += 0;
                else
                    AccDifference += Ratio / 1.5;

            }

            //
            // Right side
            //
            for (i = Right; i > (int)iLineBase && i > 0; i--)
            {
                Ratio = iFFTValues[i - 1] / iFFTValues[i];
                if (Ratio > 2)
                    AccDifference += 1.5;
                else if (Ratio < 0.5)
                    AccDifference += 0;
                else
                    AccDifference += Ratio / 1.5;

            }

            AccDifference /= Right - Left;

            return AccDifference;
        }



    }
    public class MeasReadingTimeWaveform
    {
        public uint TimeWaveformPoints = 0;

        public bool IsLoaded { get { return MeasReading.IsLoaded; } }

        private MeasReading _MeasReading = null;
        public MeasReading MeasReading { get { return _MeasReading; } }

        public AnalystConnection Connection { get { return MeasReading.Connection; } }

        private string _ReadingHeaderText;
        public string ReadingHeaderText { get { return (MeasReading.IsLoaded ? _ReadingHeaderText : string.Empty); } }

        private float[] _ReadingHeaderBlob;
        public float[] ReadingHeaderBlob { get { return (MeasReading.IsLoaded ? _ReadingHeaderBlob : null); } }

        private float _Factor;

        public uint _Points;
        public uint Points { get { return (MeasReading.IsLoaded ? _Points : 0); } }

        private float[] _ReadingData;
        public float[] ReadingData { get { return (MeasReading.IsLoaded ? _ReadingData : null); } }

        private float _Speed;
        public float Speed { get { return (MeasReading.IsLoaded ? _Speed : float.NaN); } }

        private float _DurationSeconds;
        public float DurationSeconds { get { return (MeasReading.IsLoaded ? _DurationSeconds : float.NaN); } }


        public MeasReadingTimeWaveform() { }
        public MeasReadingTimeWaveform(Measurement Measurement)
        {
            _MeasReading = new Analyst.MeasReading(Measurement);
            _Load();
        }
        public MeasReadingTimeWaveform(AnalystConnection AnalystConnection, uint MeasId)
        {
            _MeasReading = new Analyst.MeasReading(new Measurement(AnalystConnection, MeasId));
            _Load();
        }
        public MeasReadingTimeWaveform(Point Point)
        {
            //uint MeasId = Point.Connection.SQLtoUInt("max(MeasId)", "Measurement", "PointId=" + Point.PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Point.Connection, "SKFCM_ASMD_FFT").ToString());
            uint MeasId = Point.Connection.SQLtoUInt("max(DataDTG)", "Measurement", "PointId=" + Point.PointId.ToString() + " and ReadingType=" + Registration.RegistrationId(Point.Connection, "SKFCM_ASMD_Time").ToString());
            _MeasReading = new Analyst.MeasReading(new Measurement(Point));
            _Load();
        }


        private bool _Load()
        {
            //_IsLoaded = false;

            if (MeasReading.Measurement.IsLoaded)
            {
                DataTable MeasReadingRows = MeasReading.Measurement.Connection.DataTable("ReadingId, ReadingHeader, ReadingData", "MeasReading", "MeasId=" + MeasReading.Measurement.MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(MeasReading.Measurement.Connection, "SKFCM_ASMD_Time").ToString());
                if (MeasReadingRows.Rows.Count > 0)
                {
                    MeasReading.MeasReadingRow = MeasReadingRows.Rows[0];
                    MeasReading.ReadingId = Convert.ToUInt32(MeasReading.MeasReadingRow["ReadingId"]); //SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_FFT").ToString());
                    _ReadingHeaderText = Connection.BLOBtoString(MeasReading.MeasReadingRow["ReadingHeader"]); //select ReadingHeader from " + Owner + "MeasReading where ReadingId=" + FFTReadingId.ToString());
                    _Factor = GenericTools.Hex8ToValue(GenericTools.ReverseHex(_ReadingHeaderText.Substring(48, 16)));
                    List<float> HeaderCreator = new List<float>();
                    for (int i = 0; i < (int)(_ReadingHeaderText.Length / 16); i++)
                    {
                        HeaderCreator.Add(GenericTools.Hex8ToValue(GenericTools.ReverseHex(_ReadingHeaderText.Substring(16 * i, 16))));
                    }
                    _ReadingHeaderBlob = HeaderCreator.ToArray<float>();// Connection.GetBLOBtoArray(MeasReading.MeasReadingRow["ReadingHeader"], 1);
                    _ReadingData = Connection.GetBLOBtoArray(MeasReading.MeasReadingRow["ReadingData"], _Factor);
                    _Points = Convert.ToUInt32(_ReadingData.Length); //SQLtoInt("(((dbms_lob.getLength(ReadingData))/2)-1)", "MeasReading", "ReadingId=" + FFTReadingId.ToString());
                    _Speed = _ReadingHeaderBlob[2];
                    _DurationSeconds = _ReadingHeaderBlob[6];
                    // FFTAverages = 
                    //_IsLoaded = true;
                }
            }

            return IsLoaded;
        }


        public double CrestFactor()
        {
            double[] data = Array.ConvertAll(_ReadingData, x => (double)x); ;

            int iStart = 0;
            int iStop = data.Length - 1;
            int iLength = data.Length;
            double[] iValues = data;

            int i;
            double Result = 0;
            double Peak = 0;

            if (iStop == 0)
                iStop = iLength - 1;


            if (iStart > iStop || iStart > iLength || iStop >= iLength)
                return Result;

            for (i = iStart; i < iStop; i++)
            {
                if (Math.Abs(iValues[i]) > Peak)
                    Peak = Math.Abs(iValues[i]);

                Result += iValues[i] * iValues[i];  //square
            }

            Result /= iStop - iStart + 1;         // mean
            Result = Math.Sqrt(Result);   // Result now holds RMS

            Result = Peak / Result;

            return Result;
        }


        //----------------------------------------------------------------------------
        // DESC:
        //    This function calculates the standard statistical moments of a data set.
        //    The iStart and iStop values will limit the calculations to be operated
        //    to just the set starting with index iStart and ending with iStop (inclusive)
        //    Specify 0 for both values will assume iStart == 0 and iStop == iLenght -1;
        // PARAMS:
        //    iValues - array of doubles being the input values
        //    iLength - length of iValues[]
        //    iStart - Starting index within iValues[], typical "0" when starting at
        //          the first element.
        //    iStop - Last element to include in calculations, typical iLength-1 when
        //          including the last possible element in iValues[].
        //    oAverage - double, output average of all the values in iValues[]
        //    oVariance - double pointer, statistical variance of 
        //    oAvgDev - doubel pointer, average deviation
        //    oStdDev - double pointer, standard deviation
        //    oSkew - double pointer, standard statistical method 'skew'
        //    oKurtosis - double pointer, standard statistical method 'Kurtosis'
        //    oPosPeak - double pointer, true most positive peak (de-trended)
        //    oNegPeak - double pointer, true most negative peak (de-trended)
        //    oRMS - double pointer, true RMS value (de-trended)
        //----------------------------------------------------------------------------
        public bool pCalculate_Moments
           (
              int iStart,
              int iStop,
              out double outAverage,
              out double outVariance,
              out double outAvgDev,
              out double outStdDev,
              out double outSkew,
              out double outKurtosis,
              out double outPosPeak,
              out double outNegPeak,
              out double outRMS
           )
        {
            bool Result = false;

            double[] iValues = Array.ConvertAll(_ReadingData, x => (double)x);

            int iLength = iValues.Length;

            if (iStop == 0)
                iStop = iLength - 1;

            int i, Length = iStop - iStart + 1;

            outAverage = 0;
            outVariance = 0;
            outAvgDev = 0;
            outStdDev = 0;
            outSkew = 0;
            outKurtosis = 0;
            outPosPeak = 0;
            outNegPeak = 0;
            outRMS = 0;


            //
            // Calculate the average
            //
            for (i = iStart; i < iStop; i++)
                outAverage += iValues[i];
            outAverage /= Length;

            double Temp1, Temp2;
            for (i = iStart; i < iStop; i++)
            {
                Temp1 = iValues[i] - outAverage;

                outAvgDev += Math.Abs(Temp1);

                Temp2 = Temp1 * Temp1;      // Temp2 = squared
                outVariance += Temp2;

                Temp2 *= Temp1;             // Temp2 = cubed
                outSkew += Temp2;

                Temp2 *= Temp1;             // Temp 2 = quad
                outKurtosis += Temp2;

                if (iValues[i] > outPosPeak)
                    outPosPeak = iValues[i];

                if (iValues[i] < outNegPeak)
                    outNegPeak = iValues[i];

                outRMS += iValues[i] * iValues[i];
            }

            outAvgDev /= Length;
            outVariance /= Length - 1;
            outStdDev = Math.Sqrt(outVariance);

            if (outVariance != 0.0)
            {
                outSkew /= (Length * outStdDev * outStdDev * outStdDev);
                outKurtosis /= (Length * outVariance * outVariance) - 3.0;
            }

            outRMS /= Length;
            outRMS = Math.Sqrt(outRMS);

            return true;
        }



        public float GetAmplitude(uint Line) { return (IsLoaded ? _ReadingData[Line] : float.NaN); }
        //public float GetAmplitude() { return _Factor; }

        public float[] BandData() { return BandData(1, _Points); }
        public float[] BandData(uint Point) { return BandData(Point, Point); }
        public float[] BandData(uint FirstPoint, uint LastPoint)
        {
            if (!IsLoaded) return null;

            float[] ReturnValue = new float[2];
            float ReturnOverall = 0;
            float ReturnPeak = 0;

            uint Line1 = Math.Min(FirstPoint, LastPoint);
            uint Line2 = Math.Max(FirstPoint, LastPoint);

            if (Line1 < 1) Line1 = 1;
            if (Line2 > TimeWaveformPoints) Line2 = TimeWaveformPoints;

            if (Line1 == Line2)
            {
                ReturnValue[0] = _ReadingData[Line1 - 1];
                ReturnValue[1] = ReturnValue[0];
            }
            else
            {
                for (uint i = Line1; i <= Line2; i++)
                {
                    ReturnOverall += (float)Math.Pow(_ReadingData[i - 1], 2);
                    ReturnPeak = Math.Max(ReturnPeak, _ReadingData[i - 1]);
                }
                ReturnValue[0] = (float)Math.Sqrt(ReturnOverall / 1);
                ReturnValue[1] = ReturnPeak;
            }
            return ReturnValue;
        }
    }
    public class MeasReadingInspection
    {
        public bool IsLoaded { get { return (ReadingId > 0); } }
        public bool _IsLoaded = false;

        private AnalystConnection _Connection = null;
        public AnalystConnection Connection { get { return (_Connection != null ? _Connection : null); } }

        private DataRow _MeasReadingRow;
        public DataRow MeasReadingRow { get { return _MeasReadingRow; } }

        public uint MeasId { get { return (IsLoaded ? Convert.ToUInt32(_MeasReadingRow["MeasId"]) : 0); } }
        public uint PointId { get { return (IsLoaded ? Convert.ToUInt32(_MeasReadingRow["PointId"]) : 0); } }
        public uint _MeasId = 0;
        public uint _PointId = 0;

        private uint _ReadingId = 0;
        public uint ReadingId { get { return _ReadingId; } }

        public bool Inspection { get { return !(MeasReadingRow == null); } }

        public uint InspectionResult = 0;
        public string[] InspectionOptions;

        //public MeasReadingInspection MeasReadingInspection;
        public MeasReadingInspection(Measurement Measurement)
        {
            _Connection = Measurement.Connection;
            DataTable MeasReadingRows = Connection.DataTable("*", "MeasReading", "MeasId=" + Measurement.MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection"));
            if (MeasReadingRows.Rows.Count > 0)
            {

                _MeasReadingRow = MeasReadingRows.Rows[0];
                _ReadingId = Convert.ToUInt32(MeasReadingRows.Rows[0]["ReadingId"]);
                _IsLoaded = IsLoaded;
                _PointId = PointId;
                _MeasId = MeasId;
                InspectionOptions = GetInspection();

            }
            else
                _MeasReadingRow = null;
        }

        public MeasReadingInspection(AnalystConnection AnalystConnection, uint MeasId)
        {
            _Connection = AnalystConnection;

            DataTable MeasReadingRows = Connection.DataTable("*", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection"));
            if (MeasReadingRows.Rows.Count > 0)
            {
                _MeasReadingRow = MeasReadingRows.Rows[0];
                _ReadingId = Convert.ToUInt32(MeasReadingRows.Rows[0]["ReadingId"]);
                _IsLoaded = IsLoaded;
                _PointId = PointId;
                _MeasId = MeasId;
                InspectionOptions = GetInspection();
            }
            else
                _MeasReadingRow = null;
        }


        public string[] GetInspection_new(uint PointId)
        {
            GenericTools.DebugMsg("GetInspection(): Starting...");

            string[] ReturnValue = new string[5];

            if (Inspection)
            {
                uint InspectionAlrmId = Connection.SQLtoUInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString());
                if (InspectionAlrmId < 1) InspectionAlrmId = Connection.SQLtoUInt("InspectionAlrmId", "InspectionAlarm", "ElementId=" + PointId.ToString());
                ReturnValue[0] = Connection.SQLtoString("InspectionText1", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
                ReturnValue[1] = Connection.SQLtoString("InspectionText2", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
                ReturnValue[2] = Connection.SQLtoString("InspectionText3", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
                ReturnValue[3] = Connection.SQLtoString("InspectionText4", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
                ReturnValue[4] = Connection.SQLtoString("InspectionText5", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
            }


            GenericTools.DebugMsg("GetInspection(): " + ReturnValue);

            return ReturnValue;
        }



        public string[] GetInspection()
        {
            GenericTools.DebugMsg("GetInspection(): Starting...");

            string[] ReturnValue = new string[5];

            if (Inspection)
            {
                uint InspectionAlrmId = Connection.SQLtoUInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString());
                if (InspectionAlrmId < 1) InspectionAlrmId = Connection.SQLtoUInt("InspectionAlrmId", "InspectionAlarm", "ElementId=" + PointId.ToString());
                ReturnValue[0] = Connection.SQLtoString("InspectionText1", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
                ReturnValue[1] = Connection.SQLtoString("InspectionText2", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
                ReturnValue[2] = Connection.SQLtoString("InspectionText3", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
                ReturnValue[3] = Connection.SQLtoString("InspectionText4", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
                ReturnValue[4] = Connection.SQLtoString("InspectionText5", "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString());
            }


            GenericTools.DebugMsg("GetInspection(): " + ReturnValue);

            return ReturnValue;
        }

        public string ResultInspection()
        {
            string ReturnValue = string.Empty;
            ReturnValue = ResultInspection(Connection, MeasId);
            return ReturnValue;
        }

        public DataTable PointSon(AnalystConnection Connec, uint Point)
        {

            return Connection.DataTable("select ELEMENTID, TREEELEM.* from point, TREEELEM where TREEELEMID=ELEMENTID AND FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Conditional_Point') and VALUESTRING='" + Point + "'");

        }

        public string InpectionReadResult(uint InspectionResult, string Delimiter, Point Point)
        {
            Debug.Log("Verificando Inspeção do ponto: " + Point.Name);
            string ReturnValue = "";

            DataTable TreeElemdResult = PointSon(Connection, Point.TreeElemId); //Connection.SQLtoUInt("SELECT TREEELEMID FROM TREEELEM WHERE NAME LIKE 'MA %" + InspectionOptions[i].Substring(0,4) + "%' AND PARENTID=" + Point.TreeElem.TreeElemId);
            if (TreeElemdResult.Rows.Count > 0)
            {
                Debug.Log("Existem " + TreeElemdResult.Rows.Count + " filhos do ponto: " + Point.Name);
                foreach (DataRow dr1 in TreeElemdResult.Rows)
                {
                    Point pt_2 = new Point(Connection, uint.Parse(dr1["TREEELEMID"].ToString()));
                    Debug.Log("Filho: " + pt_2.Name);
                    if (pt_2.LastMeas.Inspection)
                    {
                        for (int i = 0; i < 5; i++)
                        {
                            if ((InspectionResult & (int)Math.Pow(2, i)) > 0)
                            {

                                string _ResultInspection_2 = Connection.SQLtoString("INSPECTIONTEXT" + (i + 1), "InspectionAlarm", "ELEMENTID=" + Point.PointId);
                                Debug.Log("Inspeção: " + (i + 1) + " resultado: " + _ResultInspection_2);
                                if (_ResultInspection_2 == "")
                                {
                                    if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + ", ";

                                    uint AlarmeID_2 = Connection.SQLtoUInt("SELECT  MAX([ALARMID]) AS ALARMEID  FROM [ALARMASSIGN]  where ELEMENTID=" + Point.PointId + " and TYPE=(SELECT REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASAT_Inspection')");
                                    Debug.Log("Buscando alarm em: " + AlarmeID_2);
                                    string _tempReturnValue = Connection.SQLtoString("INSPECTIONTEXT" + (i + 1), "InspectionAlarm", "INSPECTIONALRMID=" + AlarmeID_2);
                                    ReturnValue += _tempReturnValue;
                                    Debug.Log("Econtrado:  " + ReturnValue);
                                    if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + Delimiter + " ";
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + ", ";
                                    ReturnValue = ReturnValue + _ResultInspection_2;
                                }
                            }
                        }
                        Debug.Log("Verificando ultima medição do: " + pt_2.Name);
                        Measurement _Measurement_2 = new Measurement(Connection, pt_2.LastMeas.MeasId);
                        int InspectionReadingId_2 = Connection.SQLtoInt("ReadingId", "MeasReading", "MeasId=" + pt_2.LastMeas.MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection").ToString());
                        uint InspectionResult_2 = Connection.SQLtoUInt("ExDWordVal1", "MeasReading", "ReadingId=" + InspectionReadingId_2.ToString());

                        Debug.Log("Resultado da inspeção encontrada: " + InspectionResult_2);
                        if (InspectionResult_2 > 0)
                        {
                            if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + Delimiter + " ";
                            Debug.Log("Verifica resultado da inspeção em: " + pt_2.Name);
                            ReturnValue += InpectionReadResult(InspectionResult_2, ",", pt_2);

                        }
                    }
                }
            }
            else
            {
                Debug.Log("Não Existem filhos do ponto: " + Point.Name);
                for (int i = 0; i < 5; i++)
                {
                    if ((InspectionResult & (int)Math.Pow(2, i)) > 0)
                    {

                        Debug.Log("Encontrada ispeção: " + (i + 1));
                        string _ResultInspection_2 = Connection.SQLtoString("INSPECTIONTEXT" + (i + 1), "InspectionAlarm", "ELEMENTID=" + Point.PointId);
                        if (_ResultInspection_2 == "")
                        {
                            if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + Delimiter + " ";
                            Debug.Log("Verifica resultado da inspeção em: " + Point.Name);
                            uint AlarmeID_2 = Connection.SQLtoUInt("SELECT  MAX([ALARMID]) AS ALARMEID  FROM [ALARMASSIGN]  where ELEMENTID=" + Point.PointId + " and TYPE=(SELECT REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASAT_Inspection')");
                            Debug.Log("Buscando alarm em: " + AlarmeID_2);
                            ReturnValue += Connection.SQLtoString("INSPECTIONTEXT" + (i + 1), "InspectionAlarm", "INSPECTIONALRMID=" + AlarmeID_2);
                            Debug.Log("Resultado da inspeção encontrada: " + ReturnValue);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + ", ";
                            ReturnValue = ReturnValue + _ResultInspection_2;
                        }
                    }
                }
            }
            return ReturnValue;
        }
        public string ResultInspection_New()
        {

            Measurement _Measurement = new Measurement(Connection, MeasId);
            Point pt = new Point(Connection, _Measurement.PointId);

            int InspectionReadingId = Connection.SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Inspection").ToString());
            InspectionResult = (uint)Connection.SQLtoInt("ExDWordVal1", "MeasReading", "ReadingId=" + InspectionReadingId.ToString());
            //string _ResultInspection = Connection.SQLtoString("INSPECTIONTEXT" + InspectionResult, "InspectionAlarm", "ELEMENTID=" + _Measurement.PointId);

            string Resultado = InpectionReadResult(InspectionResult, ";", pt);

            if (Resultado == "")
            {

                for (int i = 0; i < 5; i++)
                {
                    if ((InspectionResult & (int)Math.Pow(2, i)) > 0)
                    {
                        if (!string.IsNullOrEmpty(Resultado)) Resultado = Resultado + ", ";

                        Debug.Log("Encontrada ispeção: " + (i + 1));
                        string _ResultInspection_2 = Connection.SQLtoString("INSPECTIONTEXT" + (i + 1), "InspectionAlarm", "ELEMENTID=" + pt.PointId);
                        if (_ResultInspection_2 == "")
                        {
                            Debug.Log("Verifica resultado da inspeção em: " + pt.Name);
                            uint AlarmeID_2 = Connection.SQLtoUInt("SELECT  MAX([ALARMID]) AS ALARMEID  FROM [ALARMASSIGN]  where ELEMENTID=" + pt.PointId + " and TYPE=(SELECT REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASAT_Inspection')");
                            Debug.Log("Buscando alarm em: " + AlarmeID_2);
                            string temp_Resultado = Connection.SQLtoString("INSPECTIONTEXT" + (i + 1), "InspectionAlarm", "INSPECTIONALRMID=" + AlarmeID_2);
                            Resultado = Resultado + temp_Resultado;

                            Debug.Log("Resultado da inspeção encontrada: " + Resultado);
                        }
                    }
                }
            }

            return Resultado.Replace("; ;", "; ").Replace(", , ", ", ");
        }



        public string ResultInspection(AnalystConnection AnalystConnection, uint MeasId)
        {
            Measurement _Measurement = new Measurement(AnalystConnection, MeasId);
            int InspectionReadingId = AnalystConnection.SQLtoInt("ReadingId", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(AnalystConnection, "SKFCM_ASMD_Inspection").ToString());
            InspectionResult = (uint)AnalystConnection.SQLtoInt("ExDWordVal1", "MeasReading", "ReadingId=" + InspectionReadingId.ToString());
            string _ResultInspection = AnalystConnection.SQLtoString("INSPECTIONTEXT" + InspectionResult, "InspectionAlarm", "ELEMENTID=" + _Measurement.PointId);

            if (_ResultInspection == "")
            {
                uint AlarmeID = AnalystConnection.SQLtoUInt("SELECT  MAX([ALARMID]) AS ALARMEID  FROM [skfuser].[skfuser1].[ALARMASSIGN]  where ELEMENTID=" + _Measurement.PointId + " and TYPE=(SELECT REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASAT_Inspection')");
                _ResultInspection = AnalystConnection.SQLtoString("INSPECTIONTEXT" + InspectionResult, "InspectionAlarm", "INSPECTIONALRMID=" + AlarmeID);
                uint TreeElemdResult = AnalystConnection.SQLtoUInt("SELECT TREEELEMID FROM TREEELEM WHERE NAME LIKE '%" + _ResultInspection + "%' AND PARENTID=" + _Measurement.Point.TreeElem.ParentId);
                if (TreeElemdResult > 0)
                {
                    Point pt_2 = new Point(AnalystConnection, TreeElemdResult);
                    Measurement _Measurement_2 = new Measurement(AnalystConnection, pt_2.LastMeas.MeasId);
                    int InspectionReadingId_2 = AnalystConnection.SQLtoInt("ReadingId", "MeasReading", "MeasId=" + pt_2.LastMeas.MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(AnalystConnection, "SKFCM_ASMD_Inspection").ToString());
                    InspectionResult = (uint)AnalystConnection.SQLtoInt("ExDWordVal1", "MeasReading", "ReadingId=" + InspectionReadingId_2.ToString());
                    string _ResultInspection_2 = AnalystConnection.SQLtoString("INSPECTIONTEXT" + InspectionResult, "InspectionAlarm", "ELEMENTID=" + _Measurement_2.PointId);

                    if (_ResultInspection_2 == "")
                    {
                        uint AlarmeID_2 = AnalystConnection.SQLtoUInt("SELECT  MAX([ALARMID]) AS ALARMEID  FROM [skfuser].[skfuser1].[ALARMASSIGN]  where ELEMENTID=" + pt_2.PointId + " and TYPE=(SELECT REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASAT_Inspection')");
                        _ResultInspection += ": " + AnalystConnection.SQLtoString("INSPECTIONTEXT" + InspectionResult, "InspectionAlarm", "INSPECTIONALRMID=" + AlarmeID_2);

                    }
                }


            }

            return _ResultInspection;
        }


        public string GetInspectionString()
        {
            return GetInspectionString(GetInspection(), ";");
        }
        public string GetInspectionString(string Delimiter)
        {
            GenericTools.DebugMsg("GetInspectionString(" + Delimiter + "): Starting...");

            string ReturnValue = string.Empty;

            if (Inspection)
            {
                ReturnValue = ReturnValue[0] + Delimiter + ReturnValue[1] + Delimiter + ReturnValue[2] + Delimiter + ReturnValue[3] + Delimiter + ReturnValue[4] + Delimiter;
            }

            GenericTools.DebugMsg("GetInspectionString(" + Delimiter + "): " + ReturnValue);

            return ReturnValue;
        }
        public string GetInspectionString(string[] GetInspection, string Delimiter)
        {
            GenericTools.DebugMsg("GetInspectionString(" + Delimiter + "): Starting...");

            string ReturnValue = string.Empty;

            if (Inspection)
            {
                ReturnValue = GetInspection[0] + Delimiter + GetInspection[1] + Delimiter + GetInspection[2] + Delimiter + GetInspection[3] + Delimiter + GetInspection[4] + Delimiter;
                ReturnValue = ReturnValue.Replace(";;", "");
            }

            GenericTools.DebugMsg("GetInspectionString(" + Delimiter + "): " + ReturnValue);

            return ReturnValue;
        }


        public string[] GetInspectionResult()
        {
            GenericTools.DebugMsg("GetInspectionResult(): Starting...");

            string[] ReturnValue = new string[5];

            for (int i = 0; i < 5; i++)
                if ((InspectionResult & (int)Math.Pow(2, i)) > 0) ReturnValue[i] = InspectionOptions[i];

            GenericTools.DebugMsg("GetInspectionResult(): " + ReturnValue);

            return ReturnValue;
        }


        public string GetInspectionResultString()
        {
            ResultInspection();
            return GetInspectionResultString(";");
        }
        public string GetInspectionResultString(string Delimiter)
        {
            ResultInspection();
            GenericTools.DebugMsg("GetInspectionResultString(" + Delimiter + "): Starting...");

            string ReturnValue = string.Empty;

            if (Inspection)
                for (int i = 0; i < 5; i++)
                    if ((InspectionResult & (int)Math.Pow(2, i)) > 0)
                    {
                        if (!string.IsNullOrEmpty(ReturnValue)) ReturnValue = ReturnValue + Delimiter;
                        ReturnValue = ReturnValue + InspectionOptions[i];
                    }

            GenericTools.DebugMsg("GetInspectionResultString(" + Delimiter + "): " + ReturnValue);

            return ReturnValue;
        }
        public uint InspectionResultAlarmFlag
        {
            get
            {
                GenericTools.DebugMsg("GetInspectionResultAlarmFlag(): Starting...");

                uint ReturnValue = 0;

                uint InspectionAlrmId = Connection.SQLtoUInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString());
                if (InspectionAlrmId < 1) InspectionAlrmId = Connection.SQLtoUInt("InspectionAlrmId", "InspectionAlarm", "ElementId=" + PointId.ToString());

                for (int i = 0; i < 5; i++)
                    if ((InspectionResult & (uint)Math.Pow(2, i)) > 0) ReturnValue = Convert.ToUInt32(Math.Max(ReturnValue, Connection.SQLtoInt("AlarmLevel" + (i + 1).ToString(), "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString())));

                GenericTools.DebugMsg("GetInspectionResultAlarmFlag(): " + ReturnValue);

                return ReturnValue;
            }
        }
        /**
        public int GetInspectionResultAlarmFlag(Int32 PointId, int tInspecionResult)
        {
            GenericTools.DebugMsg("GetInspectionResultAlarmFlag(): Starting...");

            int ReturnValue = 0;

            Int32 InspectionAlrmId = Connection.SQLtoInt("AlarmId", "AlarmAssign", "ElementId=" + PointId.ToString());
            if (InspectionAlrmId < 1) InspectionAlrmId = Connection.SQLtoInt("InspectionAlrmId", "InspectionAlarm", "ElementId=" + PointId.ToString());

            for (int i = 0; i < 5; i++)
                if ((tInspecionResult & (int)Math.Pow(2, i)) > 0) ReturnValue = Math.Max(ReturnValue, Connection.SQLtoInt("AlarmLevel" + (i + 1).ToString(), "InspectionAlarm", "InspectionAlrmId=" + InspectionAlrmId.ToString()));

            GenericTools.DebugMsg("GetInspectionResultAlarmFlag(): " + ReturnValue);

            return ReturnValue;
        }
        **/

    }
    public class MeasReadingMCD
    {
        public bool IsLoaded { get { return (ReadingId > 0); } }
        private AnalystConnection _Connection = null;
        public AnalystConnection Connection { get { return (_Connection != null ? _Connection : null); } }
        private DataRow _MeasReadingRow;
        public DataRow MeasReadingRow { get { return _MeasReadingRow; } }
        public uint MeasId { get { return (IsLoaded ? Convert.ToUInt32(_MeasReadingRow["MeasId"]) : 0); } }
        public uint PointId { get { return (IsLoaded ? Convert.ToUInt32(_MeasReadingRow["PointId"]) : 0); } }
        private Point _Point;
        public Point Point
        {
            get
            {
                if (IsLoaded)
                {
                    if (_Point == null) _Point = new Point(Connection, PointId);
                }
                else
                    _Point = null;

                return _Point;
            }
        }
        private uint _ReadingId = 0;
        public uint ReadingId { get { return _ReadingId; } }
        public bool MCD { get { return !(MeasReadingRow == null); } }
        public MCDParam Envelope { get { return _LoadMCDParams(MCDParamType.Envelope); } }
        public MCDParam Velocity { get { return _LoadMCDParams(MCDParamType.Velocity); } }
        public MCDParam Temperature { get { return _LoadMCDParams(MCDParamType.Temperature); } }
        private MCDParam _Envelope;
        private MCDParam _Velocity;
        private MCDParam _Temperature;
        private MCDParam _LoadMCDParams(MCDParamType Type)
        {
            if (IsLoaded)
            {
                _Envelope.Type = MCDParamType.Envelope;
                _Velocity.Type = MCDParamType.Velocity;
                _Temperature.Type = MCDParamType.Temperature;

                if (Point.FullScaleUnit.IndexOf(',') > 0)
                {
                    _Envelope.FullScaleUnit = Point.FullScaleUnit.Split(',')[0].Trim();
                    _Velocity.FullScaleUnit = Point.FullScaleUnit.Split(',')[1].Trim();
                    _Temperature.FullScaleUnit = Point.FullScaleUnit.Split(',')[2].Trim();
                }
                else
                {
                    _Envelope.FullScaleUnit = null;
                    _Velocity.FullScaleUnit = null;
                    _Temperature.FullScaleUnit = null;
                }
            }
            else
            {
                _Envelope = null;
                _Velocity = null;
                _Temperature = null;
            }
            switch (Type)
            {
                case MCDParamType.Envelope: return _Envelope;
                case MCDParamType.Velocity: return _Velocity;
                case MCDParamType.Temperature: return _Temperature;
            }
            return null;
        }
        public MCDMeasAlarm MeasAlarm;

        public MeasReadingMCD() { }
        public MeasReadingMCD(Measurement Measurement)
        {
            _Connection = Measurement.Connection;
            DataTable MeasReadingRows = Connection.DataTable("*", "MeasReading", "MeasId=" + Measurement.MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_MCD"));
            if (MeasReadingRows.Rows.Count > 0)
                _MeasReadingRow = MeasReadingRows.Rows[0];
            else
                _MeasReadingRow = null;
        }

        public MeasReadingMCD(AnalystConnection AnalystConnection, uint MeasId)
        {
            _Connection = AnalystConnection;
            DataTable MeasReadingRows = Connection.DataTable("*", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_MCD"));
            if (MeasReadingRows.Rows.Count > 0)
                _MeasReadingRow = MeasReadingRows.Rows[0];
            else
                _MeasReadingRow = null;
        }
    }
    public class MeasReadingMeasurementParameter : MeasReading { }
    public class MeasReadingMotorCurrent : MeasReading { }
    public class MeasReadingNonCollection : MeasReading { }
    public class MeasReadingOverall
    {
        public bool IsLoaded { get { return (MeasReadingRow != null); } }

        private AnalystConnection _Connection = null;
        public AnalystConnection Connection { get { return (_Connection != null ? _Connection : null); } }

        private DataRow _MeasReadingRow;
        public DataRow MeasReadingRow { get { return _MeasReadingRow; } }

        public uint MeasId { get { return (IsLoaded ? Convert.ToUInt32(_MeasReadingRow["MeasId"]) : 0); } }
        public uint PointId { get { return (IsLoaded ? Convert.ToUInt32(_MeasReadingRow["PointId"]) : 0); } }
        private Point _Point;
        public Point Point
        {
            get
            {
                if (IsLoaded)
                {
                    if (_Point == null) _Point = new Point(Connection, PointId);
                }
                else
                    _Point = null;

                return _Point;
            }
        }

        private uint _ReadingId = 0;
        public uint ReadingId
        {
            get
            {
                if (IsLoaded & (_ReadingId == 0)) _ReadingId = Convert.ToUInt32(_MeasReadingRow["ReadingId"]);
                return _ReadingId;
            }
        }

        public bool Overall { get { return !(MeasReadingRow == null); } }

        public MeasReadingOverall() { }
        public MeasReadingOverall(Measurement Measurement)
        {
            _Connection = Measurement.Connection;
            DataTable MeasReadingRows = Connection.DataTable("*", "MeasReading", "MeasId=" + Measurement.MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall"));
            if (MeasReadingRows.Rows.Count > 0)
                _MeasReadingRow = MeasReadingRows.Rows[0];
            else
                _MeasReadingRow = null;
        }
        public MeasReadingOverall(AnalystConnection AnalystConnection, uint MeasId)
        {
            _Connection = AnalystConnection;
            DataTable MeasReadingRows = Connection.DataTable("*", "MeasReading", "MeasId=" + MeasId.ToString() + " and ReadingType=" + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall"));
            if (MeasReadingRows.Rows.Count > 0)
                _MeasReadingRow = MeasReadingRows.Rows[0];
            else
                _MeasReadingRow = null;
        }

        public float OverallValue { get { return (IsLoaded ? Convert.ToSingle(_MeasReadingRow["OverallValue"]) : 0); } }

        private static uint GetOverall_LastMeasId = 0;
        private static float GetOverall_LastResult = 0;
        /**
        public float GetOverall(uint nMeasId)
        {
            if (nMeasId == GetOverall_LastMeasId) return GetOverall_LastResult;
            float tmpGetOverall_LastResult = Connection.SQLtoFloat("OverallValue", "MeasReading", "ReadingType = " + Registration.RegistrationId(Connection, "SKFCM_ASMD_Overall").ToString() + " and MeasId=" + nMeasId.ToString());
            if (float.IsNaN(tmpGetOverall_LastResult) | float.IsInfinity(tmpGetOverall_LastResult)) return tmpGetOverall_LastResult;
            GetOverall_LastMeasId = nMeasId;
            GetOverall_LastResult = tmpGetOverall_LastResult;
            _OverallValue = tmpGetOverall_LastResult;
            return tmpGetOverall_LastResult;
        }

        public float GetOverall()
        {
            return GetOverall(this.MeasId);
        }

        public void AnMeasOverall()
        {
            _OverallValue = float.NaN;
            _FullScaleUnit = string.Empty;
        }
         **/

        /**public string LoadFullScaleUnit(AnConnection Analyst, Int32 PointId)
        {
            GenericTools.DebugMsg("LoadFullScaleUnit(" + PointId.ToString() + "): Starting...");

            string ReturnValue = string.Empty;

            try
            {
                if (Analyst.IsConnected)
                {
                    DataTable Point = Analyst.DataTable("ValueString", "Point", "ElementId=" + PointId.ToString() + " and FieldId=" + Analyst.GetRegistrationId("SKFCM_ASPF_Full_Scale_Unit").ToString());

                    if (Point.Rows.Count > 0)
                        ReturnValue = Point.Rows[0]["ValueString"].ToString();
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("LoadFullScaleUnit(" + PointId.ToString() + ") error: " + ex.Message);
            }

            FullScaleUnit = ReturnValue;

            GenericTools.DebugMsg("LoadFullScaleUnit(" + PointId.ToString() + "): " + ReturnValue);
            return ReturnValue;
        }
        **/
    }
    public class MeasReadingParametricDigital : MeasReading { }
    public class MeasReadingPhase : MeasReading { }
    public class MeasReadingTime : MeasReading { }
    public class MCDParam
    {
        public float OverallValue;
        public string FullScaleUnit;
        public MCDParamType Type;
    }
    public class AnUser
    {
        /// <summary>
        /// Return true if user data was loaded from database
        /// </summary>
        public bool IsLoaded { get { return _IsLoaded; } }
        private bool _IsLoaded = false;

        /// <summary>
        /// Name used to login into SKF @ptitude Analyst
        /// </summary>
        public string LoginName
        {
            get
            {
                return _LoginName;
            }
            set
            {
                _LoginName = value;
                _IsLoaded = false;
            }
        }
        private string _LoginName = string.Empty;

        /// <summary>
        /// User's password
        /// </summary>
        public string PassWd { set { _PassWd = value; } }
        private string _PassWd;
        private string _PassFromTable = string.Empty;

        /// <summary>
        /// Internal id for user into UserTbl table
        /// </summary>
        public Int32 UserId
        {
            get
            {
                return _UserId;
            }
            set
            {
                _UserId = value;
                _IsLoaded = false;
            }
        }
        private Int32 _UserId = 0;

        /// <summary>
        /// Customized Access Definition Id
        /// </summary>
        public int AccessDefId
        {
            get
            {
                if (IsLoaded && SystemAccessDefId == SystemDefId.Custom)
                    return _AccessDefId;
                else
                    return 0;
            }
        }
        private int _AccessDefId = 0;

        /// <summary>
        /// System Access Definition Id
        /// </summary>
        public SystemDefId SystemAccessDefId
        {
            get
            {
                if (IsLoaded)
                    return _SystemAccessDefId;
                else
                    return SystemDefId.None;
            }
        }
        private SystemDefId _SystemAccessDefId = SystemDefId.None;

        /// <summary>
        /// Initialize a empty user
        /// </summary>
        public AnUser()
        {
            _IsLoaded = false;
            LoginName = string.Empty;
            PassWd = string.Empty;
            UserId = 0;
            _AccessDefId = 0;
            _SystemAccessDefId = 0;
        }
        /// <summary>
        /// Initialize user loading parameters
        /// </summary>
        /// <param name="iUserId">User unique ID</param>
        public AnUser(AnConnection Analyst, Int32 iUserId)
        {
            Load(Analyst, iUserId);
        }
        public AnUser(AnConnection Analyst, string sLoginName)
        {
            Load(Analyst, sLoginName);
        }
        public Int32 Load(AnConnection Analyst)
        {
            if (IsLoaded) return _UserId;

            if (_UserId > 0) return Load(Analyst, _UserId);

            return Load(Analyst, _LoginName);
        }
        public Int32 Load(AnConnection Analyst, string sLoginName)
        {
            return Load(Analyst, Analyst.SQLtoInt("UserId", "UserTbl", "upper(LoginName)='" + sLoginName.ToUpper() + "'"));
        }
        public Int32 Load(AnConnection Analyst, Int32 iUserId)
        {
            _IsLoaded = false;
            _UserId = 0;
            _LoginName = string.Empty;
            _PassWd = string.Empty;
            _PassFromTable = string.Empty;
            _AccessDefId = 0;
            _SystemAccessDefId = SystemDefId.None;

            if ((!Analyst.IsConnected) || (iUserId <= 0)) return 0;

            DataTable UserTbl = Analyst.DataTable("LoginName, PassWd, AccessDefId, SystemAccessDefId", "UserTbl", "UserId=" + iUserId.ToString());

            if (UserTbl.Rows.Count > 0)
            {
                _UserId = iUserId;
                _LoginName = UserTbl.Rows[0]["LoginName"].ToString();
                _PassFromTable = UserTbl.Rows[0]["PassWd"].ToString();
                _AccessDefId = GenericTools.StrToInt(UserTbl.Rows[0]["AccessDefId"].ToString());

                _SystemAccessDefId = (SystemDefId)GenericTools.StrToInt(UserTbl.Rows[0]["SystemAccessDefId"].ToString());


                _IsLoaded = true;
            }
            return _UserId;
        }
    }
    public class WorkSpace
    {
        #region Properties
        public TreeElem TreeElem;
        public uint HierarchyId { get { return TreeElem.TreeElemId; } }

        private AnalystConnection _Connection = null;
        public AnalystConnection Connection
        {
            get
            {
                if (TreeElem == null)
                    return _Connection;
                else
                    return TreeElem.Connection;
            }
            set
            {
                _Connection = value;
            }
        }

        public uint TblSetId { get { return TreeElem.TblSetId; } }
        public Analyst.AlarmFlags AlarmFlags { get { return TreeElem.AlarmFlags; } }
        public string Name { get { return TreeElem.Name; } }
        #endregion
        #region WorkSpace Class Constructor
        public WorkSpace(AnalystConnection AnalystConection)
        {
            Connection = AnalystConection;
        }
        public WorkSpace(AnalystConnection AnalystConection, uint WorkSpaceId)
        {
            TreeElem = new Analyst.TreeElem(AnalystConection, WorkSpaceId);
        }
        public WorkSpace(AnalystConnection AnalystConection, uint HierarchyId, string WorkSpaceName)
        {
            GenericTools.DebugMsg("WorkSpace(" + HierarchyId.ToString() + ", " + WorkSpaceName + "): Starting...");
            try
            {
                if (AnalystConection.SQLtoUInt("TreeElemId", "TreeElem", "TblSetId=" + HierarchyId + " and HierarchyType=3 and ContainerType=1 and ParentId!=2147000000 and Name='" + WorkSpaceName + "'") < 1)
                {
                    GenericTools.DebugMsg("WorkSpace(" + HierarchyId.ToString() + ", " + WorkSpaceName + "): Inserting into TreeElem table");

                    List<TableColumn> Columns = new List<TableColumn>();

                    Columns.Add(new TableColumn("TreeElemId", 0));
                    Columns.Add(new TableColumn("HierarchyId", 0));
                    Columns.Add(new TableColumn("BranchLevel", 0));
                    Columns.Add(new TableColumn("SlotNumber", 1));
                    Columns.Add(new TableColumn("TblSetId", HierarchyId));
                    Columns.Add(new TableColumn("Name", WorkSpaceName));
                    Columns.Add(new TableColumn("ContainerType", 1));
                    Columns.Add(new TableColumn("Description", WorkSpaceName));
                    Columns.Add(new TableColumn("ElementEnable", 1));
                    Columns.Add(new TableColumn("ParentEnable", 0));
                    Columns.Add(new TableColumn("HierarchyType", 3));
                    Columns.Add(new TableColumn("AlarmFlags", 0));
                    Columns.Add(new TableColumn("ParentId", 0));
                    Columns.Add(new TableColumn("ParentRefId", 0));
                    Columns.Add(new TableColumn("Good", 0));
                    Columns.Add(new TableColumn("Alert", 0));
                    Columns.Add(new TableColumn("Danger", 0));
                    Columns.Add(new TableColumn("Overdue", 0));
                    Columns.Add(new TableColumn("ChannelEnable", 1));

                    uint TreeElemIdTMP = Convert.ToUInt32(AnalystConection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_seq"));

                    if (TreeElemIdTMP > 0)
                    {
                        GenericTools.DebugMsg("WorkSpace(" + HierarchyId.ToString() + ", " + WorkSpaceName + "): Inserting into Workspace table");

                        Columns.Clear();

                        Columns.Add(new TableColumn("ElementId", TreeElemIdTMP));
                        Columns.Add(new TableColumn("FilterId", 0));
                        Columns.Add(new TableColumn("LastViewed", ""));
                        Columns.Add(new TableColumn("LastFiltered", ""));
                        Columns.Add(new TableColumn("FilterUpdateOption", 0));
                        Columns.Add(new TableColumn("Summary", ""));
                        Columns.Add(new TableColumn("FilterSourceId", 0));
                        Columns.Add(new TableColumn("FilterRootId", 0));
                        Columns.Add(new TableColumn("FilterSelectionId", 0));
                        Columns.Add(new TableColumn("KeepHierStruct", 0));

                        TreeElemIdTMP = Convert.ToUInt32(AnalystConection.SQLInsert("Workspace", Columns));
                    }
                }

                uint ElementId_Workspace = AnalystConection.SQLtoUInt("ELEMENTID", "WORKSPACE", "ELEMENTID IN (SELECT TREEELEMID FROM TREEELEM WHERE TblSetId=" + HierarchyId + " and HierarchyType=3 and ContainerType=1 and ParentId!=2147000000 and Name='" + WorkSpaceName + "')");

                if (ElementId_Workspace < 1)
                {
                    uint TreeElemId = AnalystConection.SQLtoUInt("SELECT TREEELEMID FROM TREEELEM WHERE TblSetId=" + HierarchyId + " and HierarchyType=3 and ContainerType=1 and ParentId!=2147000000 and Name='" + WorkSpaceName + "'");
                    GenericTools.DebugMsg("WorkSpace(" + HierarchyId.ToString() + ", " + WorkSpaceName + "): Inserting into Workspace table");

                    List<TableColumn> Columns = new List<TableColumn>();
                    Columns.Clear();
                    Columns.Add(new TableColumn("ElementId", TreeElemId));
                    Columns.Add(new TableColumn("FilterId", 0));
                    Columns.Add(new TableColumn("LastViewed", ""));
                    Columns.Add(new TableColumn("LastFiltered", ""));
                    Columns.Add(new TableColumn("FilterUpdateOption", 0));
                    Columns.Add(new TableColumn("Summary", ""));
                    Columns.Add(new TableColumn("FilterSourceId", 0));
                    Columns.Add(new TableColumn("FilterRootId", 0));
                    Columns.Add(new TableColumn("FilterSelectionId", 0));
                    Columns.Add(new TableColumn("KeepHierStruct", 0));

                    TreeElemId = Convert.ToUInt32(AnalystConection.SQLInsert("Workspace", Columns));
                }

                TreeElem = new Analyst.TreeElem(AnalystConection, AnalystConection.SQLtoUInt("TreeElemId", "TreeElem", "TblSetId=" + HierarchyId + " and HierarchyType=3 and ContainerType=1 and ParentId!=2147000000 and Name='" + WorkSpaceName + "'"));
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("WorkSpace(" + HierarchyId.ToString() + ", " + WorkSpaceName + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("WorkSpace(" + HierarchyId.ToString() + ", " + WorkSpaceName + "): " + TreeElem.TreeElemId.ToString());
        }
        #endregion
        #region Functions
        public bool DeleteWorkSpace(uint WorkSpaceId, uint TblSetId)
        {
            bool ChangeTreeElem = Connection.SQLExec("UPDATE TREEELEM SET PARENTID=2147000000 WHERE (HIERARCHYID=" + WorkSpaceId + " OR TREEELEMID=" + WorkSpaceId + ") AND TBLSETID=" + TblSetId);
            bool ChangeWorkspace = Connection.SQLExec("DELETE FROM WORKSPACE WHERE ELEMENTID=" + WorkSpaceId);

            return (ChangeTreeElem && ChangeWorkspace);
        }
        public bool DeleteWorkSpace(string WorkSpaceName, uint TblSetId)
        {
            GenericTools.DebugMsg("Deleting Workspace: " + WorkSpaceName + " tblsetid: " + TblSetId);
            bool ChangeTreeElem = Connection.SQLExec("UPDATE TREEELEM SET PARENTID=2147000000 WHERE (TREEELEMID=(SELECT TREEELEMID FROM TREEELEM WHERE NAME='" + WorkSpaceName + "' AND TBLSETID=" + TblSetId + " AND PARENTID!=2147000000) OR HIERARCHYID=(SELECT TREEELEMID FROM TREEELEM WHERE NAME='" + WorkSpaceName + "' AND TBLSETID=" + TblSetId + " AND PARENTID!=2147000000)) AND TBLSETID=" + TblSetId);
            bool ChangeWorkspace = Connection.SQLExec("DELETE FROM WORKSPACE WHERE ELEMENTID=(SELECT TREEELEMID FROM TREEELEM WHERE NAME='" + WorkSpaceName + "' AND TBLSETID=" + TblSetId + " AND PARENTID!=2147000000)");

            GenericTools.DebugMsg("Deleting Workspace, ChangeTreeElem: " + ChangeTreeElem + " ChangeWorkspace: " + ChangeWorkspace);

            return (ChangeTreeElem && ChangeWorkspace);
        }

        public void SortSlotNumer(Set set, string Type = "ASC")
        {
            DataTable dt = Connection.DataTable("SELECT * FROM TREEELEM WHERE PARENTID=" + set.TreeElem.TreeElemId + " ORDER BY NAME " + Type);

            int i = 1;
            foreach (DataRow dr in dt.Rows)
            {
                Connection.SQLExec("UPDATE TREEELEM SET SLOTNUMBER=" + i + " WHERE TREEELEMID=" + dr["TREEELEMID"].ToString());
                i++;
            }
        }

        public Machine FindMachine(uint MachineId) { return FindMachine(new Machine(Connection, MachineId)); }
        public Machine FindMachine(Machine Machine)
        {
            GenericTools.DebugMsg("Workspace.FindMachine(" + Machine.TreeElemId.ToString() + "): Starting...");

            Machine ReturnValue = null;

            try
            {
                uint MachineId = Connection.SQLtoUInt("min(ParentId)", "TreeElem", "HierarchyId=" + HierarchyId.ToString() + " and ParentId!=2147000000 and ParentRefId=" + Machine.TreeElemId.ToString());
                if ((MachineId != uint.MinValue) & (MachineId > 0))
                {
                    ReturnValue = new Machine(Connection, MachineId);
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Workspace.FindMachine(" + Machine.TreeElemId.ToString() + ") error: " + ex.Message);
            }

            if (ReturnValue == null)
                GenericTools.DebugMsg("Workspace.FindMachine(" + Machine.TreeElemId.ToString() + "): null");
            else
                GenericTools.DebugMsg("Workspace.FindMachine(" + Machine.TreeElemId.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }
        public Set FindSet(Set ParentSet, string SetName)
        {
            GenericTools.DebugMsg("FindSet(" + ParentSet.TreeElem.TreeElemId.ToString() + ", " + SetName + "): Starting...");

            try
            {
                uint SetId = Connection.SQLtoUInt("TreeElemId", "TreeElem", "ParentId=" + ParentSet.TreeElem.TreeElemId.ToString() + " and Name='" + SetName + "'");
                if ((SetId != uint.MinValue) & (SetId > 0))
                {
                    GenericTools.DebugMsg("FindSet(" + ParentSet.TreeElem.TreeElemId.ToString() + ", " + SetName + "): " + SetId.ToString());
                    return new Set(Connection, SetId);
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("FindSet(" + ParentSet.TreeElem.TreeElemId.ToString() + ", " + SetName + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("FindSet(" + ParentSet.TreeElem.TreeElemId.ToString() + ", " + SetName + "): null");

            return null;
        }

        public Machine AddMachine(uint MachineId, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null) { return AddMachine(HierarchyId, MachineId, SlotNumber, IncludePoints, NewName); }
        public Machine AddMachine(Machine Machine, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null) { return AddMachine(Machine.TreeElemId, SlotNumber, IncludePoints, NewName); }
        public Machine AddMachine(uint SetId, Machine Machine, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null) { return AddMachine(SetId, Machine.TreeElemId, SlotNumber, IncludePoints, NewName); }
        public Machine AddMachine(Set Set, Machine Machine, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null) { return AddMachine(Set.TreeElem.TreeElemId, Machine.TreeElemId, SlotNumber, IncludePoints, NewName); }
        public Machine AddMachine(uint SetId, uint MachineId, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null)
        {
            GenericTools.DebugMsg("AddMachine() Starting...");
            Machine ReturnValue = null;

            try
            {
                uint ParentId = HierarchyId;
                string _NewName = null;

                TreeElem TreeElemSet = new TreeElem(Connection, SetId);
                if (TreeElemSet.HierarchyId == HierarchyId)
                    ParentId = TreeElemSet.TreeElemId;
                TreeElemSet = new TreeElem(Connection, ParentId);

                TreeElem TreeElemMachine = new TreeElem(Connection, MachineId);
                if (TreeElemMachine.HierarchyType == HierarchyType.Hierarchy)
                {


                    if (NewName != null)
                    {
                        _NewName = NewName;
                    }
                    else
                    {
                        _NewName = TreeElemMachine.Name;
                    }

                    List<TableColumn> Columns = new List<TableColumn>();

                    Columns.Clear();

                    Columns.Add(new TableColumn("TreeElemId", 0));
                    Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                    Columns.Add(new TableColumn("BranchLevel", (TreeElemSet.BranchLevel + 1)));
                    Columns.Add(new TableColumn("SlotNumber", (SlotNumber > 0 ? SlotNumber : (Connection.SQLtoUInt("max(SlotNumber)", "TreeElem", "ParentId=" + TreeElemSet.TreeElemId.ToString()) + 1))));
                    Columns.Add(new TableColumn("TblSetId", TreeElemSet.TblSetId));
                    Columns.Add(new TableColumn("Name", _NewName));
                    Columns.Add(new TableColumn("ContainerType", 3));
                    Columns.Add(new TableColumn("Description", TreeElemMachine.Description));
                    Columns.Add(new TableColumn("ElementEnable", (TreeElemMachine.ElementEnable ? 1 : 0)));
                    Columns.Add(new TableColumn("ParentEnable", (TreeElemMachine.ParentEnable ? 1 : 0)));
                    Columns.Add(new TableColumn("HierarchyType", 3));
                    Columns.Add(new TableColumn("AlarmFlags", (uint)TreeElemMachine.AlarmFlags));
                    Columns.Add(new TableColumn("ParentId", TreeElemSet.TreeElemId));
                    Columns.Add(new TableColumn("ParentRefId", TreeElemSet.TreeElemId));
                    Columns.Add(new TableColumn("ReferenceId", 0));
                    Columns.Add(new TableColumn("Good", 0));
                    Columns.Add(new TableColumn("Alert", 0));
                    Columns.Add(new TableColumn("Danger", 0));
                    Columns.Add(new TableColumn("Overdue", (TreeElemMachine.Overdue ? 1 : 0)));
                    Columns.Add(new TableColumn("ChannelEnable", (TreeElemMachine.ChannelEnable ? 1 : 0)));

                    ReturnValue = new Machine(Connection, Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_Seq")));

                    if (ReturnValue != null)
                    {
                        Connection.SQLExec("delete from " + Connection.Owner + "GroupTbl where ElementId=" + ReturnValue.TreeElemId.ToString());
                        Connection.SQLExec("insert into " + Connection.Owner + "GroupTbl (ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority) select " + ReturnValue.TreeElemId.ToString() + " as ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority from " + Connection.Owner + "GroupTbl where ElementId=" + MachineId.ToString());
                        if (IncludePoints)
                        {
                            foreach (TreeElem TreeElemPoint in TreeElemMachine.Child)
                                AddPoint(ReturnValue, new Point(TreeElemPoint));

                            ReturnValue.TreeElem.CalcAlarm();
                        }

                        /**
                        // Calculate alarms based on childs
                        Connection.SQLExec("update " + Connection.Owner + "TreeElem set " +
                            "AlarmFlags=(select max(AlarmFlags) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElemId.ToString() + ")," +
                            "Good=(select sum(Good) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElemId.ToString() + ")," +
                            "Alert=(select sum(Alert) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElemId.ToString() + ")," +
                            "Danger=(select sum(Danger) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElemId.ToString() + ")" +
                            " where TreeElemId=" + ReturnValue.TreeElemId.ToString());

                        // Calculate parent's alarms
                        Connection.SQLExec("update " + Connection.Owner + "TreeElem set " +
                            "AlarmFlags=(select max(AlarmFlags) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.ParentId.ToString() + ")," +
                            "Good=(select sum(Good) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.ParentId.ToString() + ")," +
                            "Alert=(select sum(Alert) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.ParentId.ToString() + ")," +
                            "Danger=(select sum(Danger) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.ParentId.ToString() + ")" +
                            " where TreeElemId=" + ReturnValue.TreeElem.ParentId.ToString());
                        **/
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("AddMachine() error: " + ex.Message);
            }

            GenericTools.DebugMsg("AddMachine(): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public Set AddSet(uint SetId) { return AddSet(new Set(Connection, SetId)); }
        public Set AddSet(Set Set) { return AddSet(HierarchyId, Set); }
        public Set AddSet(uint ParentSetId, Set Set) { return AddSet(new Set(Connection, ParentSetId), Set); }
        public Set AddSet(uint ParentSetId, uint SetId) { return AddSet(new Set(Connection, ParentSetId), new Set(Connection, SetId)); }
        public Set AddSet(Set ParentSet, string NewSetName, string NewSetDescription = "", uint SlotNumber = 0)
        {
            GenericTools.DebugMsg("AddSet(" + NewSetName + "): Starting...");
            Set ReturnValue = FindSet(ParentSet, NewSetName);

            if (ReturnValue == null)
            {
                try
                {
                    if ((ParentSet.TreeElem.HierarchyType == HierarchyType.Workspace) & ((ParentSet.TreeElem.ContainerType == ContainerType.Root) | (ParentSet.TreeElem.ContainerType == ContainerType.Set)))
                    {
                        List<TableColumn> Columns = new List<TableColumn>();

                        Columns.Clear();

                        Columns.Add(new TableColumn("TreeElemId", 0));
                        Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                        Columns.Add(new TableColumn("BranchLevel", (ParentSet.TreeElem.BranchLevel + 1)));
                        Columns.Add(new TableColumn("SlotNumber", (SlotNumber > 0 ? SlotNumber : (Connection.SQLtoUInt("max(SlotNumber)", "TreeElem", "ParentId=" + ParentSet.TreeElem.TreeElemId.ToString()) + 1))));
                        Columns.Add(new TableColumn("TblSetId", ParentSet.TreeElem.TblSetId));
                        Columns.Add(new TableColumn("Name", NewSetName));
                        Columns.Add(new TableColumn("ContainerType", 2));
                        Columns.Add(new TableColumn("Description", NewSetDescription));
                        Columns.Add(new TableColumn("ElementEnable", 1));
                        Columns.Add(new TableColumn("ParentEnable", ((!ParentSet.TreeElem.ElementEnable) | (ParentSet.TreeElem.ParentEnable) ? 1 : 0)));
                        Columns.Add(new TableColumn("HierarchyType", 3));
                        Columns.Add(new TableColumn("AlarmFlags", 0));
                        Columns.Add(new TableColumn("ParentId", ParentSet.TreeElem.TreeElemId));
                        Columns.Add(new TableColumn("ParentRefId", ParentSet.TreeElem.TreeElemId));
                        Columns.Add(new TableColumn("ReferenceId", 0));
                        Columns.Add(new TableColumn("Good", 0));
                        Columns.Add(new TableColumn("Alert", 0));
                        Columns.Add(new TableColumn("Danger", 0));
                        Columns.Add(new TableColumn("Overdue", 0));
                        Columns.Add(new TableColumn("ChannelEnable", 1));

                        ReturnValue = new Set(Connection, Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_Seq")));

                        if (ReturnValue != null)
                        {
                            Connection.SQLExec("delete from " + Connection.Owner + "GroupTbl where ElementId=" + ReturnValue.TreeElem.TreeElemId.ToString());
                            Connection.SQLExec("insert into " + Connection.Owner + "GroupTbl (ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority) select " + ReturnValue.TreeElem.TreeElemId.ToString() + " as ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority from " + Connection.Owner + "GroupTbl where ElementId=" + ParentSet.TreeElem.TreeElemId.ToString());
                        }

                    }
                }
                catch (Exception ex)
                {
                    GenericTools.DebugMsg("AddSet(" + NewSetName + ") error: " + ex.Message);
                }
            }
            GenericTools.DebugMsg("AddSet(" + NewSetName + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        public Set AddSet(Set ParentSet, Set Set, bool WithChild = true, uint SlotNumber = 0)
        {
            GenericTools.DebugMsg("AddSet(): Starting...");
            Set ReturnValue = FindSet(ParentSet, Set.TreeElem.Name);

            if (ReturnValue == null)
            {
                try
                {
                    if ((ParentSet.TreeElem.HierarchyType == HierarchyType.Workspace) & (Set.TreeElem.HierarchyType == HierarchyType.Hierarchy))
                    {
                        List<TableColumn> Columns = new List<TableColumn>();

                        Columns.Clear();

                        Columns.Add(new TableColumn("TreeElemId", 0));
                        Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                        Columns.Add(new TableColumn("BranchLevel", (ParentSet.TreeElem.BranchLevel + 1)));
                        Columns.Add(new TableColumn("SlotNumber", (SlotNumber > 0 ? SlotNumber : (Connection.SQLtoUInt("max(SlotNumber)", "TreeElem", "ParentId=" + ParentSet.TreeElem.TreeElemId.ToString()) + 1))));
                        Columns.Add(new TableColumn("TblSetId", ParentSet.TreeElem.TblSetId));
                        Columns.Add(new TableColumn("Name", Set.TreeElem.Name));
                        Columns.Add(new TableColumn("ContainerType", 2));
                        Columns.Add(new TableColumn("Description", Set.TreeElem.Description));
                        Columns.Add(new TableColumn("ElementEnable", (ParentSet.TreeElem.ElementEnable ? 1 : 0)));
                        Columns.Add(new TableColumn("ParentEnable", (ParentSet.TreeElem.ParentEnable ? 1 : 0)));
                        Columns.Add(new TableColumn("HierarchyType", 3));
                        Columns.Add(new TableColumn("AlarmFlags", 0));
                        Columns.Add(new TableColumn("ParentId", ParentSet.TreeElem.TreeElemId));
                        Columns.Add(new TableColumn("ParentRefId", ParentSet.TreeElem.TreeElemId));
                        Columns.Add(new TableColumn("ReferenceId", 0));
                        Columns.Add(new TableColumn("Good", 0));
                        Columns.Add(new TableColumn("Alert", 0));
                        Columns.Add(new TableColumn("Danger", 0));
                        Columns.Add(new TableColumn("Overdue", (Set.TreeElem.Overdue ? 1 : 0)));
                        Columns.Add(new TableColumn("ChannelEnable", (Set.TreeElem.ChannelEnable ? 1 : 0)));

                        ReturnValue = new Set(Connection, Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_Seq")));

                        if (ReturnValue != null)
                        {
                            Connection.SQLExec("delete from " + Connection.Owner + "GroupTbl where ElementId=" + ReturnValue.TreeElem.TreeElemId.ToString());
                            Connection.SQLExec("insert into " + Connection.Owner + "GroupTbl (ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority) select " + ReturnValue.TreeElem.TreeElemId.ToString() + " as ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority from " + Connection.Owner + "GroupTbl where ElementId=" + Set.TreeElem.TreeElemId.ToString());
                            if (WithChild)
                            {
                                foreach (TreeElem Child in Set.TreeElem.Child)
                                    switch (Child.ContainerType)
                                    {
                                        case ContainerType.Set:
                                            AddSet(ReturnValue.TreeElem.TreeElemId, new Set(Child));
                                            break;

                                        case ContainerType.Machine:
                                            AddMachine(ReturnValue.TreeElem.TreeElemId, new Machine(Child));
                                            break;
                                    }
                                // Update alarms with child
                                Connection.SQLExec("update " + Connection.Owner + "TreeElem set " +
                    "AlarmFlags=(select max(AlarmFlags) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Good=(select sum(Good) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Alert=(select sum(Alert) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Danger=(select sum(Danger) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")" +
                    " where TreeElemId=" + ReturnValue.TreeElem.TreeElemId.ToString());
                                // Update parent's alarms
                                Connection.SQLExec("update " + Connection.Owner + "TreeElem set " +
                    "AlarmFlags=(select max(AlarmFlags) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Good=(select sum(Good) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Alert=(select sum(Alert) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Danger=(select sum(Danger) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")" +
                    " where TreeElemId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString());
                                //if (ReturnValue.TreeElem.Child.Count < 1) ReturnValue.TreeElem.Delete();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    GenericTools.DebugMsg("AddSet() error: " + ex.Message);
                }
            }
            GenericTools.DebugMsg("AddSet(): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public Point AddPoint(uint MachineId, uint PointId) { return AddPoint(new Machine(Connection, MachineId), new Point(Connection, PointId)); }
        public Point AddPoint(Machine Machine, uint PointId) { return AddPoint(Machine, new Point(Connection, PointId)); }
        public Point AddPoint(uint MachineId, Point Point) { return AddPoint(new Machine(Connection, MachineId), Point); }
        public Point AddPoint(Machine Machine, Point Point)
        {
            Point ReturnValue = null;

            if ((Machine.TreeElem.HierarchyType == HierarchyType.Workspace) & (Point.TreeElem.HierarchyType == HierarchyType.Hierarchy))
            {
                List<TableColumn> Columns = new List<TableColumn>();

                Columns.Clear();

                Columns.Add(new TableColumn("TreeElemId", 0));
                Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                Columns.Add(new TableColumn("BranchLevel", (Machine.TreeElem.BranchLevel + 1)));
                Columns.Add(new TableColumn("SlotNumber", (Connection.SQLtoUInt("max(SlotNumber)", "TreeElem", "ParentId=" + Machine.TreeElem.TreeElemId.ToString()) + 1)));
                Columns.Add(new TableColumn("TblSetId", Machine.TreeElem.TblSetId));
                Columns.Add(new TableColumn("Name", Point.TreeElem.Name));
                Columns.Add(new TableColumn("ContainerType", 4));
                Columns.Add(new TableColumn("Description", Point.TreeElem.Description));
                Columns.Add(new TableColumn("ElementEnable", (Point.TreeElem.ElementEnable ? 1 : 0)));
                Columns.Add(new TableColumn("ParentEnable", (Point.TreeElem.ParentEnable ? 1 : 0)));
                Columns.Add(new TableColumn("HierarchyType", 3));
                Columns.Add(new TableColumn("AlarmFlags", (uint)Point.TreeElem.AlarmFlags));
                Columns.Add(new TableColumn("ParentId", Machine.TreeElem.TreeElemId));
                Columns.Add(new TableColumn("ParentRefId", Point.TreeElem.Parent.TreeElemId));
                Columns.Add(new TableColumn("ReferenceId", Point.TreeElem.TreeElemId));
                Columns.Add(new TableColumn("Good", Point.TreeElem.Good));
                Columns.Add(new TableColumn("Alert", Point.TreeElem.Alert));
                Columns.Add(new TableColumn("Danger", Point.TreeElem.Danger));
                Columns.Add(new TableColumn("Overdue", (Point.TreeElem.Overdue ? 1 : 0)));
                Columns.Add(new TableColumn("ChannelEnable", (Point.TreeElem.ChannelEnable ? 1 : 0)));

                ReturnValue = new Point(Connection, Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_Seq")));
            }

            return ReturnValue;
        }

        public bool RemoveElement(Set Set) { return RemoveElement(Set.TreeElem); }
        public bool RemoveElement(Machine Machine) { return RemoveElement(Machine.TreeElem); }
        public bool RemoveElement(Point Point) { return RemoveElement(Point.TreeElem); }
        public bool RemoveElement(uint TreeElemId) { return RemoveElement(new TreeElem(Connection, TreeElemId)); }
        public bool RemoveElement(TreeElem TreeElem)
        {
            GenericTools.DebugMsg("Worspace.RemoveElement(" + TreeElem.TreeElemId.ToString() + ") Starting...");
            bool ReturnValue = false;

            try
            {
                //if (TreeElem.HierarchyType == HierarchyType.Workspace)
                //{

                //    if (TreeElem.HierarchyId == HierarchyId)
                //    {
                //        ReturnValue = true;
                //        if (TreeElem.ContainerType != ContainerType.Point)
                //            foreach (TreeElem ChildInstance in TreeElem.Child)
                //                ReturnValue &= RemoveElement(ChildInstance);
                //        if (TreeElem.ContainerType != ContainerType.Root)
                //        {
                //            TreeElem OriginalParent = TreeElem.Parent;
                //            ReturnValue &= TreeElem.Delete(true, true);

                //            if (OriginalParent.Child.Count == 1)
                //                Console.Write("");

                //            if (ReturnValue)
                //                if (OriginalParent.Child.Count <= 0)
                //                    ReturnValue &= RemoveElement(OriginalParent);

                //        }
                //    }
                if (TreeElem.HierarchyType == HierarchyType.Workspace)
                {
                    if (TreeElem.HierarchyId == HierarchyId)
                    {
                        ReturnValue = true;

                        if (TreeElem.ContainerType != ContainerType.Point)
                            foreach (TreeElem ChildInstance in TreeElem.Child)
                                ReturnValue &= RemoveElement(ChildInstance);

                        if (TreeElem.ContainerType != ContainerType.Root)
                        {
                            TreeElem OriginalParent = TreeElem.Parent;
                            if (OriginalParent.Child.Count == 1 || OriginalParent.Child.Count == 0)
                                Console.Write("");

                            ReturnValue &= TreeElem.Delete(true, true);
                            //if (ReturnValue)
                            if (OriginalParent.Child.Count <= 0)
                                ReturnValue &= RemoveElement(OriginalParent);

                        }
                    }
                }
                else
                {
                    if (TreeElem.ContainerType != ContainerType.Point)
                        foreach (DataRow TreeElemItem in Connection.DataTable("TreeElemId", "TreeElem", "HierarchyType=3 and HierarchyId=" + HierarchyId.ToString() + " and ParentId!=2147000000 and ParentRefId=" + TreeElem.TreeElemId.ToString()).Rows)
                            ReturnValue &= RemoveElement(Convert.ToUInt32(TreeElemItem["TreeElem"]));
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Worspace.RemoveElement(" + TreeElem.TreeElemId.ToString() + ") error: " + ex.Message);
                ReturnValue = false;
            }

            GenericTools.DebugMsg("Worspace.RemoveElement(" + TreeElem.TreeElemId.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }
        #endregion
    }

    public class Route
    {
        #region Properties
        public TreeElem TreeElem;
        public uint HierarchyId { get { return TreeElem.TreeElemId; } }

        private AnalystConnection _Connection = null;
        public AnalystConnection Connection
        {
            get
            {
                if (TreeElem == null)
                    return _Connection;
                else
                    return TreeElem.Connection;
            }
            set
            {
                _Connection = value;
            }
        }

        public uint TblSetId { get { return TreeElem.TblSetId; } }
        public Analyst.AlarmFlags AlarmFlags { get { return TreeElem.AlarmFlags; } }
        public string Name { get { return TreeElem.Name; } }
        #endregion
        #region Route Class Constructor
        public Route(AnalystConnection AnalystConection)
        {
            Connection = AnalystConection;
        }
        public Route(AnalystConnection AnalystConection, uint RouteId)
        {
            TreeElem = new Analyst.TreeElem(AnalystConection, RouteId);
        }
        public Route(AnalystConnection AnalystConection, uint HierarchyId, string RouteName)
        {
            GenericTools.DebugMsg("Route(" + HierarchyId.ToString() + ", " + RouteName + "): Starting...");
            try
            {
                if (AnalystConection.SQLtoUInt("TreeElemId", "TreeElem", "TblSetId=" + HierarchyId + " and HierarchyType=2 and ContainerType=1 and ParentId!=2147000000 and Name='" + RouteName + "'") < 1)
                {
                    GenericTools.DebugMsg("Route(" + HierarchyId.ToString() + ", " + RouteName + "): Inserting into TreeElem table");

                    List<TableColumn> Columns = new List<TableColumn>();

                    Columns.Add(new TableColumn("TreeElemId", 0));
                    Columns.Add(new TableColumn("HierarchyId", 0));
                    Columns.Add(new TableColumn("BranchLevel", 0));
                    Columns.Add(new TableColumn("SlotNumber", 1));
                    Columns.Add(new TableColumn("TblSetId", HierarchyId));
                    Columns.Add(new TableColumn("Name", RouteName));
                    Columns.Add(new TableColumn("ContainerType", 1));
                    Columns.Add(new TableColumn("Description", RouteName));
                    Columns.Add(new TableColumn("ElementEnable", 1));
                    Columns.Add(new TableColumn("ParentEnable", 0));
                    Columns.Add(new TableColumn("HierarchyType", 2));
                    Columns.Add(new TableColumn("AlarmFlags", 0));
                    Columns.Add(new TableColumn("ParentId", 0));
                    Columns.Add(new TableColumn("ParentRefId", 0));
                    Columns.Add(new TableColumn("Good", 0));
                    Columns.Add(new TableColumn("Alert", 0));
                    Columns.Add(new TableColumn("Danger", 0));
                    Columns.Add(new TableColumn("Overdue", 0));
                    Columns.Add(new TableColumn("ChannelEnable", 1));

                    uint TreeElemIdTMP = Convert.ToUInt32(AnalystConection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_seq"));

                    if (TreeElemIdTMP > 0)
                    {
                        GenericTools.DebugMsg("Route(" + HierarchyId.ToString() + ", " + RouteName + "): Inserting into Route table");

                        Columns.Clear();
                        Columns.Add(new TableColumn("Elementid", TreeElemIdTMP));
                        Columns.Add(new TableColumn("TblSetId", HierarchyId));
                        Columns.Add(new TableColumn("Instructions", ""));
                        Columns.Add(new TableColumn("NextSchedule", "20151209000000"));
                        Columns.Add(new TableColumn("ScheduleEnabled", 1));
                        Columns.Add(new TableColumn("Schedule", "2592000"));
                        Columns.Add(new TableColumn("ScheduleUnits", 1));
                        Columns.Add(new TableColumn("FirstShiftStartTime", 0));
                        Columns.Add(new TableColumn("LastUploaded", ""));
                        Columns.Add(new TableColumn("LastDownloaded", ""));
                        Columns.Add(new TableColumn("KeepDataMethod", 1));
                        Columns.Add(new TableColumn("KeepNumRecords", 1000));
                        Columns.Add(new TableColumn("KeepForTime", 31536000));
                        Columns.Add(new TableColumn("KeepForUnits", 4));
                        Columns.Add(new TableColumn("RouteType", 0));
                        Columns.Add(new TableColumn("ScheduleType", 0));
                        Columns.Add(new TableColumn("StartDts", "20090101080000"));
                        Columns.Add(new TableColumn("EndDts", "20090101080000"));
                        Columns.Add(new TableColumn("Settings", @"0\8\1\1\2\1\1\1\2\1\1\1\1\1\2\1\1"));
                        UInt32 ret = Convert.ToUInt32(AnalystConection.SQLInsert("RouteHdr", Columns));
                        TreeElemIdTMP = ret;
                    }
                }

                uint ElementId_Route = AnalystConection.SQLtoUInt("ELEMENTID", "RouteHdr", "ELEMENTID IN (SELECT TREEELEMID FROM TREEELEM WHERE TblSetId=" + HierarchyId + " and HierarchyType=2 and ContainerType=1 and ParentId!=2147000000 and Name='" + RouteName + "')");

                if (ElementId_Route < 1)
                {
                    uint TreeElemId = AnalystConection.SQLtoUInt("SELECT TREEELEMID FROM TREEELEM WHERE TblSetId=" + HierarchyId + " and HierarchyType=2 and ContainerType=1 and ParentId!=2147000000 and Name='" + RouteName + "'");
                    GenericTools.DebugMsg("Route(" + HierarchyId.ToString() + ", " + RouteName + "): Inserting into Route table");

                    List<TableColumn> Columns = new List<TableColumn>();
                    Columns.Clear();
                    Columns.Add(new TableColumn("TblSetId", HierarchyId));
                    Columns.Add(new TableColumn("Instructions", ""));
                    Columns.Add(new TableColumn("NextSchedule", "20151209000000"));
                    Columns.Add(new TableColumn("ScheduleEnabled", 1));
                    Columns.Add(new TableColumn("Schedule", "2592000"));
                    Columns.Add(new TableColumn("ScheduleUnits", 1));
                    Columns.Add(new TableColumn("FirstShiftStartTime", 0));
                    Columns.Add(new TableColumn("LastUploaded", ""));
                    Columns.Add(new TableColumn("LastDownloaded", ""));
                    Columns.Add(new TableColumn("KeepDataMethod", 1));
                    Columns.Add(new TableColumn("KeepNumRecords", 1000));
                    Columns.Add(new TableColumn("KeepForTime", 31536000));
                    Columns.Add(new TableColumn("KeepForUnits", 4));
                    Columns.Add(new TableColumn("RouteType", 0));
                    Columns.Add(new TableColumn("ScheduleType", 0));
                    Columns.Add(new TableColumn("StartDts", "20090101080000"));
                    Columns.Add(new TableColumn("EndDts", "20090101080000"));
                    Columns.Add(new TableColumn("Settings", @"0\8\1\1\2\1\1\1\2\1\1\1\1\1\2\1\1"));

                    TreeElemId = Convert.ToUInt32(AnalystConection.SQLInsert("RouteHdr", Columns));
                }

                TreeElem = new Analyst.TreeElem(AnalystConection, AnalystConection.SQLtoUInt("TreeElemId", "TreeElem", "TblSetId=" + HierarchyId + " and HierarchyType=2 and ContainerType=1 and ParentId!=2147000000 and Name='" + RouteName + "'"));
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Route(" + HierarchyId.ToString() + ", " + RouteName + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("Route(" + HierarchyId.ToString() + ", " + RouteName + "): " + TreeElem.TreeElemId.ToString());
        }
        #endregion
        #region Functions
        public bool DeleteRoute(uint RouteId, uint TblSetId)
        {
            bool ChangeTreeElem = Connection.SQLExec("UPDATE TREEELEM SET PARENTID=2147000000 WHERE (HIERARCHYID=" + RouteId + " OR TREEELEMID=" + RouteId + ") AND TBLSETID=" + TblSetId);
            bool ChangeRoute = Connection.SQLExec("DELETE FROM Route WHERE ELEMENTID=" + RouteId);

            return (ChangeTreeElem && ChangeRoute);
        }
        public bool DeleteRoute(string RouteName, uint TblSetId)
        {
            GenericTools.DebugMsg("Deleting Route: " + RouteName + " tblsetid: " + TblSetId);
            bool ChangeTreeElem = Connection.SQLExec("UPDATE TREEELEM SET PARENTID=2147000000 WHERE (TREEELEMID=(SELECT TREEELEMID FROM TREEELEM WHERE NAME='" + RouteName + "' AND TBLSETID=" + TblSetId + " AND PARENTID!=2147000000) OR HIERARCHYID=(SELECT TREEELEMID FROM TREEELEM WHERE NAME='" + RouteName + "' AND TBLSETID=" + TblSetId + " AND PARENTID!=2147000000)) AND TBLSETID=" + TblSetId);
            bool ChangeRoute = Connection.SQLExec("DELETE FROM Route WHERE ELEMENTID=(SELECT TREEELEMID FROM TREEELEM WHERE NAME='" + RouteName + "' AND TBLSETID=" + TblSetId + " AND PARENTID!=2147000000)");

            GenericTools.DebugMsg("Deleting Route, ChangeTreeElem: " + ChangeTreeElem + " ChangeRoute: " + ChangeRoute);

            return (ChangeTreeElem && ChangeRoute);
        }

        public void SortSlotNumer(Set set, string Type = "ASC")
        {
            DataTable dt = Connection.DataTable("SELECT * FROM TREEELEM WHERE PARENTID=" + set.TreeElem.TreeElemId + " ORDER BY NAME " + Type);

            int i = 1;
            foreach (DataRow dr in dt.Rows)
            {
                Connection.SQLExec("UPDATE TREEELEM SET SLOTNUMBER=" + i + " WHERE TREEELEMID=" + dr["TREEELEMID"].ToString());
                i++;
            }
        }

        public Machine FindMachine(uint MachineId) { return FindMachine(new Machine(Connection, MachineId)); }
        public Machine FindMachine(Machine Machine)
        {
            GenericTools.DebugMsg("Route.FindMachine(" + Machine.TreeElemId.ToString() + "): Starting...");

            Machine ReturnValue = null;

            try
            {
                uint MachineId = Connection.SQLtoUInt("min(ParentId)", "TreeElem", "HierarchyId=" + HierarchyId.ToString() + " and ParentId!=2147000000 and ParentRefId=" + Machine.TreeElemId.ToString());
                if ((MachineId != uint.MinValue) & (MachineId > 0))
                {
                    ReturnValue = new Machine(Connection, MachineId);
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Route.FindMachine(" + Machine.TreeElemId.ToString() + ") error: " + ex.Message);
            }

            if (ReturnValue == null)
                GenericTools.DebugMsg("Route.FindMachine(" + Machine.TreeElemId.ToString() + "): null");
            else
                GenericTools.DebugMsg("Route.FindMachine(" + Machine.TreeElemId.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }
        public Set FindSet(Set ParentSet, string SetName)
        {
            GenericTools.DebugMsg("FindSet(" + ParentSet.TreeElem.TreeElemId.ToString() + ", " + SetName + "): Starting...");

            try
            {
                uint SetId = Connection.SQLtoUInt("TreeElemId", "TreeElem", "ParentId=" + ParentSet.TreeElem.TreeElemId.ToString() + " and Name='" + SetName + "'");
                if ((SetId != uint.MinValue) & (SetId > 0))
                {
                    GenericTools.DebugMsg("FindSet(" + ParentSet.TreeElem.TreeElemId.ToString() + ", " + SetName + "): " + SetId.ToString());
                    return new Set(Connection, SetId);
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("FindSet(" + ParentSet.TreeElem.TreeElemId.ToString() + ", " + SetName + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("FindSet(" + ParentSet.TreeElem.TreeElemId.ToString() + ", " + SetName + "): null");

            return null;
        }

        public Machine AddMachine(uint MachineId, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null) { return AddMachine(HierarchyId, MachineId, SlotNumber, IncludePoints, NewName); }
        public Machine AddMachine(Machine Machine, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null) { return AddMachine(Machine.TreeElemId, SlotNumber, IncludePoints, NewName); }
        public Machine AddMachine(uint SetId, Machine Machine, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null) { return AddMachine(SetId, Machine.TreeElemId, SlotNumber, IncludePoints, NewName); }
        public Machine AddMachine(Set Set, Machine Machine, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null) { return AddMachine(Set.TreeElem.TreeElemId, Machine.TreeElemId, SlotNumber, IncludePoints, NewName); }
        public Machine AddMachine(uint SetId, uint MachineId, uint SlotNumber = 0, bool IncludePoints = true, string NewName = null)
        {
            GenericTools.DebugMsg("AddMachine() Starting...");
            Machine ReturnValue = null;

            try
            {
                uint ParentId = HierarchyId;
                string _NewName = null;

                TreeElem TreeElemSet = new TreeElem(Connection, SetId);
                if (TreeElemSet.HierarchyId == HierarchyId)
                    ParentId = TreeElemSet.TreeElemId;
                TreeElemSet = new TreeElem(Connection, ParentId);

                TreeElem TreeElemMachine = new TreeElem(Connection, MachineId);
                if (TreeElemMachine.HierarchyType == HierarchyType.Hierarchy)
                {


                    if (NewName != null)
                    {
                        _NewName = NewName;
                    }
                    else
                    {
                        _NewName = TreeElemMachine.Name;
                    }

                    List<TableColumn> Columns = new List<TableColumn>();

                    Columns.Clear();

                    Columns.Add(new TableColumn("TreeElemId", 0));
                    Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                    Columns.Add(new TableColumn("BranchLevel", (TreeElemSet.BranchLevel + 1)));
                    Columns.Add(new TableColumn("SlotNumber", (SlotNumber > 0 ? SlotNumber : (Connection.SQLtoUInt("max(SlotNumber)", "TreeElem", "ParentId=" + TreeElemSet.TreeElemId.ToString()) + 1))));
                    Columns.Add(new TableColumn("TblSetId", TreeElemSet.TblSetId));
                    Columns.Add(new TableColumn("Name", _NewName));
                    Columns.Add(new TableColumn("ContainerType", 3));
                    Columns.Add(new TableColumn("Description", TreeElemMachine.Description));
                    Columns.Add(new TableColumn("ElementEnable", (TreeElemMachine.ElementEnable ? 1 : 0)));
                    Columns.Add(new TableColumn("ParentEnable", (TreeElemMachine.ParentEnable ? 1 : 0)));
                    Columns.Add(new TableColumn("HierarchyType", 2));
                    Columns.Add(new TableColumn("AlarmFlags", (uint)TreeElemMachine.AlarmFlags));
                    Columns.Add(new TableColumn("ParentId", TreeElemSet.TreeElemId));
                    Columns.Add(new TableColumn("ParentRefId", TreeElemSet.TreeElemId));
                    Columns.Add(new TableColumn("ReferenceId", 0));
                    Columns.Add(new TableColumn("Good", 0));
                    Columns.Add(new TableColumn("Alert", 0));
                    Columns.Add(new TableColumn("Danger", 0));
                    Columns.Add(new TableColumn("Overdue", (TreeElemMachine.Overdue ? 1 : 0)));
                    Columns.Add(new TableColumn("ChannelEnable", (TreeElemMachine.ChannelEnable ? 1 : 0)));

                    ReturnValue = new Machine(Connection, Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_Seq")));

                    if (ReturnValue != null)
                    {
                        Connection.SQLExec("delete from " + Connection.Owner + "GroupTbl where ElementId=" + ReturnValue.TreeElemId.ToString());
                        Connection.SQLExec("insert into " + Connection.Owner + "GroupTbl (ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority) select " + ReturnValue.TreeElemId.ToString() + " as ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority from " + Connection.Owner + "GroupTbl where ElementId=" + MachineId.ToString());
                        if (IncludePoints)
                        {
                            foreach (TreeElem TreeElemPoint in TreeElemMachine.Child)
                                AddPoint(ReturnValue, new Point(TreeElemPoint));

                            ReturnValue.TreeElem.CalcAlarm();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("AddMachine() error: " + ex.Message);
            }

            GenericTools.DebugMsg("AddMachine(): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public Set AddSet(uint SetId) { return AddSet(new Set(Connection, SetId)); }
        public Set AddSet(Set Set) { return AddSet(HierarchyId, Set); }
        public Set AddSet(uint ParentSetId, Set Set) { return AddSet(new Set(Connection, ParentSetId), Set); }
        public Set AddSet(uint ParentSetId, uint SetId) { return AddSet(new Set(Connection, ParentSetId), new Set(Connection, SetId)); }
        public Set AddSet(Set ParentSet, string NewSetName, string NewSetDescription = "", uint SlotNumber = 0)
        {
            GenericTools.DebugMsg("AddSet(" + NewSetName + "): Starting...");
            Set ReturnValue = FindSet(ParentSet, NewSetName);

            if (ReturnValue == null)
            {
                try
                {
                    if ((ParentSet.TreeElem.HierarchyType == HierarchyType.Route) & ((ParentSet.TreeElem.ContainerType == ContainerType.Root) | (ParentSet.TreeElem.ContainerType == ContainerType.Set)))
                    {
                        List<TableColumn> Columns = new List<TableColumn>();

                        Columns.Clear();

                        Columns.Add(new TableColumn("TreeElemId", 0));
                        Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                        Columns.Add(new TableColumn("BranchLevel", (ParentSet.TreeElem.BranchLevel + 1)));
                        Columns.Add(new TableColumn("SlotNumber", (SlotNumber > 0 ? SlotNumber : (Connection.SQLtoUInt("max(SlotNumber)", "TreeElem", "ParentId=" + ParentSet.TreeElem.TreeElemId.ToString()) + 1))));
                        Columns.Add(new TableColumn("TblSetId", ParentSet.TreeElem.TblSetId));
                        Columns.Add(new TableColumn("Name", NewSetName));
                        Columns.Add(new TableColumn("ContainerType", 2));
                        Columns.Add(new TableColumn("Description", NewSetDescription));
                        Columns.Add(new TableColumn("ElementEnable", 1));
                        Columns.Add(new TableColumn("ParentEnable", ((!ParentSet.TreeElem.ElementEnable) | (ParentSet.TreeElem.ParentEnable) ? 1 : 0)));
                        Columns.Add(new TableColumn("HierarchyType", 2));
                        Columns.Add(new TableColumn("AlarmFlags", 0));
                        Columns.Add(new TableColumn("ParentId", ParentSet.TreeElem.TreeElemId));
                        Columns.Add(new TableColumn("ParentRefId", ParentSet.TreeElem.TreeElemId));
                        Columns.Add(new TableColumn("ReferenceId", 0));
                        Columns.Add(new TableColumn("Good", 0));
                        Columns.Add(new TableColumn("Alert", 0));
                        Columns.Add(new TableColumn("Danger", 0));
                        Columns.Add(new TableColumn("Overdue", 0));
                        Columns.Add(new TableColumn("ChannelEnable", 1));

                        ReturnValue = new Set(Connection, Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_Seq")));

                        if (ReturnValue != null)
                        {
                            Connection.SQLExec("delete from " + Connection.Owner + "GroupTbl where ElementId=" + ReturnValue.TreeElem.TreeElemId.ToString());
                            Connection.SQLExec("insert into " + Connection.Owner + "GroupTbl (ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority) select " + ReturnValue.TreeElem.TreeElemId.ToString() + " as ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority from " + Connection.Owner + "GroupTbl where ElementId=" + ParentSet.TreeElem.TreeElemId.ToString());
                        }

                    }
                }
                catch (Exception ex)
                {
                    GenericTools.DebugMsg("AddSet(" + NewSetName + ") error: " + ex.Message);
                }
            }
            GenericTools.DebugMsg("AddSet(" + NewSetName + "): " + ReturnValue.ToString());
            return ReturnValue;
        }
        public Set AddSet(Set ParentSet, Set Set, bool WithChild = true, uint SlotNumber = 0)
        {
            GenericTools.DebugMsg("AddSet(): Starting...");
            Set ReturnValue = FindSet(ParentSet, Set.TreeElem.Name);

            if (ReturnValue == null)
            {
                try
                {
                    if ((ParentSet.TreeElem.HierarchyType == HierarchyType.Route) & (Set.TreeElem.HierarchyType == HierarchyType.Hierarchy))
                    {
                        List<TableColumn> Columns = new List<TableColumn>();

                        Columns.Clear();

                        Columns.Add(new TableColumn("TreeElemId", 0));
                        Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                        Columns.Add(new TableColumn("BranchLevel", (ParentSet.TreeElem.BranchLevel + 1)));
                        Columns.Add(new TableColumn("SlotNumber", (SlotNumber > 0 ? SlotNumber : (Connection.SQLtoUInt("max(SlotNumber)", "TreeElem", "ParentId=" + ParentSet.TreeElem.TreeElemId.ToString()) + 1))));
                        Columns.Add(new TableColumn("TblSetId", ParentSet.TreeElem.TblSetId));
                        Columns.Add(new TableColumn("Name", Set.TreeElem.Name));
                        Columns.Add(new TableColumn("ContainerType", 2));
                        Columns.Add(new TableColumn("Description", Set.TreeElem.Description));
                        Columns.Add(new TableColumn("ElementEnable", (ParentSet.TreeElem.ElementEnable ? 1 : 0)));
                        Columns.Add(new TableColumn("ParentEnable", (ParentSet.TreeElem.ParentEnable ? 1 : 0)));
                        Columns.Add(new TableColumn("HierarchyType", 2));
                        Columns.Add(new TableColumn("AlarmFlags", 0));
                        Columns.Add(new TableColumn("ParentId", ParentSet.TreeElem.TreeElemId));
                        Columns.Add(new TableColumn("ParentRefId", ParentSet.TreeElem.TreeElemId));
                        Columns.Add(new TableColumn("ReferenceId", 0));
                        Columns.Add(new TableColumn("Good", 0));
                        Columns.Add(new TableColumn("Alert", 0));
                        Columns.Add(new TableColumn("Danger", 0));
                        Columns.Add(new TableColumn("Overdue", (Set.TreeElem.Overdue ? 1 : 0)));
                        Columns.Add(new TableColumn("ChannelEnable", (Set.TreeElem.ChannelEnable ? 1 : 0)));

                        ReturnValue = new Set(Connection, Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_Seq")));

                        if (ReturnValue != null)
                        {
                            Connection.SQLExec("delete from " + Connection.Owner + "GroupTbl where ElementId=" + ReturnValue.TreeElem.TreeElemId.ToString());
                            Connection.SQLExec("insert into " + Connection.Owner + "GroupTbl (ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority) select " + ReturnValue.TreeElem.TreeElemId.ToString() + " as ElementId, GroupTypeId, Custom1, Custom2, Custom3, Custom4, Custom5, AssetName, SegmentName, Priority from " + Connection.Owner + "GroupTbl where ElementId=" + Set.TreeElem.TreeElemId.ToString());
                            if (WithChild)
                            {
                                foreach (TreeElem Child in Set.TreeElem.Child)
                                    switch (Child.ContainerType)
                                    {
                                        case ContainerType.Set:
                                            AddSet(ReturnValue.TreeElem.TreeElemId, new Set(Child));
                                            break;

                                        case ContainerType.Machine:
                                            AddMachine(ReturnValue.TreeElem.TreeElemId, new Machine(Child));
                                            break;
                                    }
                                // Update alarms with child
                                Connection.SQLExec("update " + Connection.Owner + "TreeElem set " +
                    "AlarmFlags=(select max(AlarmFlags) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Good=(select sum(Good) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Alert=(select sum(Alert) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Danger=(select sum(Danger) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")" +
                    " where TreeElemId=" + ReturnValue.TreeElem.TreeElemId.ToString());
                                // Update parent's alarms
                                Connection.SQLExec("update " + Connection.Owner + "TreeElem set " +
                    "AlarmFlags=(select max(AlarmFlags) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Good=(select sum(Good) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Alert=(select sum(Alert) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")," +
                    "Danger=(select sum(Danger) from " + Connection.Owner + "TreeElem where ParentId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString() + ")" +
                    " where TreeElemId=" + ReturnValue.TreeElem.Parent.TreeElemId.ToString());
                                //if (ReturnValue.TreeElem.Child.Count < 1) ReturnValue.TreeElem.Delete();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    GenericTools.DebugMsg("AddSet() error: " + ex.Message);
                }
            }
            GenericTools.DebugMsg("AddSet(): " + ReturnValue.ToString());
            return ReturnValue;
        }

        public Point AddPoint(uint MachineId, uint PointId) { return AddPoint(new Machine(Connection, MachineId), new Point(Connection, PointId)); }
        public Point AddPoint(Machine Machine, uint PointId) { return AddPoint(Machine, new Point(Connection, PointId)); }
        public Point AddPoint(uint MachineId, Point Point) { return AddPoint(new Machine(Connection, MachineId), Point); }
        public Point AddPoint(Machine Machine, Point Point)
        {
            Point ReturnValue = null;

            if ((Machine.TreeElem.HierarchyType == HierarchyType.Route) & (Point.TreeElem.HierarchyType == HierarchyType.Hierarchy))
            {
                List<TableColumn> Columns = new List<TableColumn>();

                Columns.Clear();

                Columns.Add(new TableColumn("TreeElemId", 0));
                Columns.Add(new TableColumn("HierarchyId", HierarchyId));
                Columns.Add(new TableColumn("BranchLevel", (Machine.TreeElem.BranchLevel + 1)));
                Columns.Add(new TableColumn("SlotNumber", (Connection.SQLtoUInt("max(SlotNumber)", "TreeElem", "ParentId=" + Machine.TreeElem.TreeElemId.ToString()) + 1)));
                Columns.Add(new TableColumn("TblSetId", Machine.TreeElem.TblSetId));
                Columns.Add(new TableColumn("Name", Point.TreeElem.Name));
                Columns.Add(new TableColumn("ContainerType", 4));
                Columns.Add(new TableColumn("Description", Point.TreeElem.Description));
                Columns.Add(new TableColumn("ElementEnable", (Point.TreeElem.ElementEnable ? 1 : 0)));
                Columns.Add(new TableColumn("ParentEnable", (Point.TreeElem.ParentEnable ? 1 : 0)));
                Columns.Add(new TableColumn("HierarchyType", 2));
                Columns.Add(new TableColumn("AlarmFlags", (uint)Point.TreeElem.AlarmFlags));
                Columns.Add(new TableColumn("ParentId", Machine.TreeElem.TreeElemId));
                Columns.Add(new TableColumn("ParentRefId", Point.TreeElem.Parent.TreeElemId));
                Columns.Add(new TableColumn("ReferenceId", Point.TreeElem.TreeElemId));
                Columns.Add(new TableColumn("Good", Point.TreeElem.Good));
                Columns.Add(new TableColumn("Alert", Point.TreeElem.Alert));
                Columns.Add(new TableColumn("Danger", Point.TreeElem.Danger));
                Columns.Add(new TableColumn("Overdue", (Point.TreeElem.Overdue ? 1 : 0)));
                Columns.Add(new TableColumn("ChannelEnable", (Point.TreeElem.ChannelEnable ? 1 : 0)));

                ReturnValue = new Point(Connection, Convert.ToUInt32(Connection.SQLInsert("TreeElem", Columns, "TreeElemId", "TreeElemId_Seq")));
            }

            return ReturnValue;
        }

        public bool RemoveElement(Set Set) { return RemoveElement(Set.TreeElem); }
        public bool RemoveElement(Machine Machine) { return RemoveElement(Machine.TreeElem); }
        public bool RemoveElement(Point Point) { return RemoveElement(Point.TreeElem); }
        public bool RemoveElement(uint TreeElemId) { return RemoveElement(new TreeElem(Connection, TreeElemId)); }
        public bool RemoveElement(TreeElem TreeElem)
        {
            GenericTools.DebugMsg("Worspace.RemoveElement(" + TreeElem.TreeElemId.ToString() + ") Starting...");
            bool ReturnValue = false;

            try
            {

                if (TreeElem.HierarchyType == HierarchyType.Route)
                {
                    if (TreeElem.HierarchyId == HierarchyId)
                    {
                        ReturnValue = true;

                        if (TreeElem.ContainerType != ContainerType.Point)
                            foreach (TreeElem ChildInstance in TreeElem.Child)
                                ReturnValue &= RemoveElement(ChildInstance);

                        if (TreeElem.ContainerType != ContainerType.Root)
                        {
                            TreeElem OriginalParent = TreeElem.Parent;
                            if (OriginalParent.Child.Count == 1 || OriginalParent.Child.Count == 0)
                                Console.Write("");

                            ReturnValue &= TreeElem.Delete(true, true);
                            //if (ReturnValue)
                            if (OriginalParent.Child.Count <= 0)
                                ReturnValue &= RemoveElement(OriginalParent);

                        }
                    }
                }
                else
                {
                    if (TreeElem.ContainerType != ContainerType.Point)
                        foreach (DataRow TreeElemItem in Connection.DataTable("TreeElemId", "TreeElem", "HierarchyType=2 and HierarchyId=" + HierarchyId.ToString() + " and ParentId!=2147000000 and ParentRefId=" + TreeElem.TreeElemId.ToString()).Rows)
                            ReturnValue &= RemoveElement(Convert.ToUInt32(TreeElemItem["TreeElem"]));
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("Route.RemoveElement(" + TreeElem.TreeElemId.ToString() + ") error: " + ex.Message);
                ReturnValue = false;
            }

            GenericTools.DebugMsg("Route.RemoveElement(" + TreeElem.TreeElemId.ToString() + "): " + ReturnValue.ToString());

            return ReturnValue;
        }
        #endregion
    }
    public class Note
    {
        #region Properties
        private bool _IsLoaded = false;
        private AnalystConnection _Connection { get; set; }
        public AnalystConnection Connection { get { return _Connection; } set { _Connection = value; } }

        private DataRow _NoteRow;
        public uint Id { get { return (_IsLoaded ? Convert.ToUInt32(_NoteRow["NotesId"]) : 0); } }

        private uint pOwner { get; set; }
        public uint OwnerId { get { return (_IsLoaded ? Convert.ToUInt32(_NoteRow["OwnerId"]) : 0); } set { pOwner = value; } }

        public TreeElem Owner { get { return (_IsLoaded ? new TreeElem(_Connection, OwnerId) : null); } }

        private string pText { get; set; }
        public string Text { get { return (_IsLoaded ? _NoteRow["Text"].ToString() : string.Empty); } set { pText = value; } }

        private DateTime pDataDtg { get; set; }
        public DateTime DataDtg { get { return (_IsLoaded ? GenericTools.StrToDateTime(_NoteRow["DataDtg"].ToString()) : DateTime.MinValue); } set { pDataDtg = value; } }

        public uint CategoryId { get { return (_IsLoaded ? Convert.ToUInt32(_NoteRow["CategoryId"]) : 0); } }

        public NoteCategory Category
        {
            get
            {
                switch (Registration.Signature(_Connection, CategoryId))
                {
                    case "SKFCM_ASNO_UserNote": return NoteCategory.UserNote;
                    case "SKFCM_ASNO_CollectionNote": return NoteCategory.CollectionNote;
                    case "SKFCM_ASNO_NonCollectionNote": return NoteCategory.NonCollectionNote;
                    case "SKFCM_ASNO_CodedNote": return NoteCategory.CodedNote;
                    case "SKFCM_ASNO_OperatingTimeResetNote": return NoteCategory.OperatingTimeResetNote;
                    case "SKFCM_ASNO_AcknowledgeAlarmNote": return NoteCategory.AcknowledgeAlarmNote;
                    case "SKFCM_ASNO_OilAnalysisNote": return NoteCategory.OilAnalysisNote;
                    default: return NoteCategory.None;
                }
            }
        }
        #endregion
        #region Note Class Constructor
        public Note(AnalystConnection Connection)
        {
            _Connection = Connection;
        }
        public Note(AnalystConnection Connection, uint NoteId) { _Load(Connection, NoteId); }
        public Note(AnalystConnection Connection, Measurement Meas)
        {
            TreeElem Te = new TreeElem(Connection, Meas.PointId);
            _Load(Te);
        }
        public Note(AnalystConnection Connection, Measurement Meas, string DateDTG)
        {
            TreeElem Te = new TreeElem(Connection, Meas.PointId);
            uint LastNoteId = Connection.SQLtoUInt("max(NotesId)", "Notes", "OwnerId=" + Te.TreeElemId.ToString() + " and DataDtg='" + DateDTG + "'");
            _Load(Connection, LastNoteId);
        }
        public Note(TreeElem TreeElem) { _Load(TreeElem); }


        public List<SKF.RS.STB.Analyst.Note> Notes(uint TreeElemId) { return Notes(TreeElemId, DateTime.MinValue, DateTime.MaxValue); }
        public List<SKF.RS.STB.Analyst.Note> Notes(uint TreeElemId, DateTime Date)
        {
            Date = new DateTime(Date.Year, Date.Month, Date.Day, 0, 0, 0);
            return Notes(TreeElemId, Date, Date.AddDays(1));
        }
        public List<SKF.RS.STB.Analyst.Note> Notes(uint TreeElemId, DateTime StartDateTime, DateTime EndDateTime)
        {
            List<SKF.RS.STB.Analyst.Note> ReturnValue = new List<SKF.RS.STB.Analyst.Note>();

            string SQL = "";
            if (StartDateTime == EndDateTime)
            {
                SQL = "SELECT NotesId FROM Notes WHERE (OwnerId=" + TreeElemId.ToString() + " or OwnerId in (select TreeElemId from " + Connection.Owner +
                    "TreeElem where ElementEnable=1 and ParentId=" + TreeElemId.ToString() + ")) and (DataDtg = '"
                    + GenericTools.DateTime(StartDateTime) + "')";
            }
            else
            {
                SQL = "SELECT NotesId FROM Notes WHERE (OwnerId=" + TreeElemId.ToString() + " or OwnerId in (select TreeElemId from " + Connection.Owner +
                    "TreeElem where ElementEnable=1 and ParentId=" + TreeElemId.ToString() + ")) and (DataDtg between '"
                    + GenericTools.DateTime(StartDateTime) + "' and '"
                    + GenericTools.DateTime(EndDateTime) + "')";
            }
            //            DataTable TableNotes = Connection.DataTable("NotesId", "Notes", "(OwnerId=" + TreeElemId.ToString() + " or OwnerId in (select TreeElemId from " + Connection.Owner + "TreeElem where ElementEnable=1 and ParentId=" + TreeElemId.ToString() + ")) and (DataDtg between '" + GenericTools.DateTime(StartDateTime) + "' and '" + GenericTools.DateTime(EndDateTime) + "')");
            DataTable TableNotes = Connection.DataTable(SQL);
            if (TableNotes.Rows.Count > 0)
                for (int i = 0; i < TableNotes.Rows.Count; i++)
                    ReturnValue.Add(new Note(Connection, Convert.ToUInt32(TableNotes.Rows[i]["NotesId"])));

            TableNotes.Dispose();
            return ReturnValue;
        }
        #endregion

        #region Functions
        private bool _Load(TreeElem TreeElem)
        {
            string LastNoteDtg = TreeElem.Connection.SQLtoString("max(DataDtg)", "Notes", "OwnerId=" + TreeElem.TreeElemId.ToString());
            uint LastNoteId = TreeElem.Connection.SQLtoUInt("max(NotesId)", "Notes", "OwnerId=" + TreeElem.TreeElemId.ToString() + " and DataDtg='" + LastNoteDtg + "'");

            return _Load(TreeElem.Connection, LastNoteId);
        }
        private bool _Load(AnalystConnection Connection, uint NoteId)
        {
            DataTable _NoteTable = Connection.DataTable("*", "Notes", "NotesId=" + NoteId.ToString());
            if (_NoteTable.Rows.Count > 0)
            {
                _NoteRow = _NoteTable.Rows[0];
                _Connection = Connection;
                _IsLoaded = true;
            }
            return _IsLoaded;
        }


        #endregion
    }

    public class TransactionServer
    {
        public uint Port;
        public string Host;

        public string SoapGetUserId()
        {
            int retorno = 2147000000;
            try
            {
                SOAPCon.SOAPCon scon = new SOAPCon.SOAPCon();
                retorno = scon.GetTransactionUserId(Host, (int)Port, "TRANSACTIONSERVER");
            }
            catch(Exception ex)
            {

            }

            return retorno.ToString();


            /**
            WebClient Teste = new WebClient();
            Teste.BaseAddress= "http://" + Host + ":" + Port.ToString();
            Teste.Headers.Add("MethodName", "Get_Table_Sets");
            Teste.Headers.Add("InterfaceName", "SKFCM_Rpc::{7361619E-E6E6-4B5E-A23A-BB0855FBEF36}");
            Teste.Headers.Add("MessageType", "Call");
            Teste.Headers.Add("Content-Type", "text/xml-SOAP");
            Teste.QueryString.Add("iUserName", UserName);
            Teste.QueryString.Add("iPassword")
            **/

            //string ReturnValue = string.Empty;

            //WebRequest SoapQuery = WebRequest.Create("http://" + Host + ":" + Port.ToString());

            //SoapQuery.Headers.Add("MethodName", "Get_Table_Sets");
            //SoapQuery.Headers.Add("InterfaceName", "SKFCM_Rpc::{7361619E-E6E6-4B5E-A23A-BB0855FBEF36}");
            //SoapQuery.Headers.Add("MessageType", "Call");
            ////SoapQuery.Headers.Add("Content-Type", "text/xml-SOAP");

            //SoapQuery.Method = "POST";

            //string Call = "<Connect_User><iUserName>" + UserName + "</iUserName><iPassword>" + Password + "</iPassword><oUserId></oUserId><oTableSetId></oTableSetId></Connect_User>";
            //byte[] CallByteArray = Encoding.UTF8.GetBytes(Call);

            ////SoapQuery.ContentType = "text/xml-SOAP";

            //SoapQuery.ContentLength = CallByteArray.Length;

            //SoapQuery.Timeout = 50000;

            //Stream dataStream = null; 
            //try
            //{
            //    dataStream = SoapQuery.GetRequestStream();
            //}
            //catch (WebException ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

            //dataStream.Write(CallByteArray, 0, CallByteArray.Length);

            //dataStream.Close();

            //WebResponse SoapResponse = SoapQuery.GetResponse();

            //string Description = ((HttpWebResponse)SoapResponse).StatusDescription;

            //dataStream = SoapResponse.GetResponseStream();

            //StreamReader reader = new StreamReader(dataStream);

            //string responseFromServer = reader.ReadToEnd();

            //ReturnValue = responseFromServer;

            // return ReturnValue;
        }
    }
}
