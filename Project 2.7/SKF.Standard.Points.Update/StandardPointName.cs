using System;
using Dapper;
using SKF.RS.STB.Analyst;
using SKF.Standard.Points.Entity;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using TreeElem = SKF.Standard.Points.Entity.TreeElem;

namespace SKF.Standard.Points.Update
{
    public class StandardPointName
    {
        private readonly SqlConnection _connection;
        public StandardPointName(AnalystConnection cn)
        {
            _connection = new SqlConnection(cn.ConnectionString);
        }
        
        public IEnumerable<TreeElem> GetPointsByTableSetId(uint id)
        {
            return _connection.Query<TreeElem>($"SELECT DISTINCT * FROM TREEELEM WHERE CONTAINERTYPE=4 AND TBLSETID={id} and Parentid!=2147000000 and ELEMENTENABLE=1");
        }

        public List<EstimateName> ParseStandardNames(IEnumerable<TreeElem> measPoints)
        {
            return (from point in measPoints let parsedName = ParseStandardName(point.NAME) 
                select new EstimateName() {PointId = point.TREEELEMID, NewName = parsedName, OldName = point.NAME,
                    HierarchyId = point.HIERARCHYID, ReferenceId = point.REFERENCEID}).ToList();
        }

        private string ParseStandardName(string str)
        {
            var angularOrientationEst = "";
            var measurementType = "";
            var machineSide = "";

            // handle ACC, ENV, VEL, E3 special cases and remove it
            measurementType = GetMeasurementType(str.ToUpper(), out str);

            string trimmedStr = TrimName(str);
            

            var numberMatch = GetAvailableNumberInString(trimmedStr);

            if (!numberMatch.Success) return string.Empty;

            // make the string always start with the number
            trimmedStr = trimmedStr.Substring(numberMatch.Index);
          
            // after having the string start always with 0, the bearing number is always on 0 start index
            var bearingNumberIndex = 0;

            var bearingNumber = trimmedStr[0].ToString() != string.Empty 
                ? int.Parse(numberMatch.Value).ToString("00") 
                : numberMatch.Value;

            if (bearingNumber!=string.Empty)
            {
                machineSide = GetMachineSide(bearingNumber, GetCode(str, MachineSide.EstimationCodes));
            }

            // if the number is the end of the string
            if (trimmedStr.Length == 1) return $"{bearingNumber} {machineSide}";

            //End of string
            // but still need to check what next to the number
            angularOrientationEst = trimmedStr.Length == bearingNumberIndex+1
                ? trimmedStr[bearingNumberIndex].ToString()
                : trimmedStr[bearingNumberIndex+1].ToString();

            // check if the angularOrientation selected one is valid
            angularOrientationEst = AngularOrientation.EstimationCodes.Exists(x => 
                x.Equals(angularOrientationEst))
                ? AngularOrientation.EstimationCodes.Single(x => x.Equals(angularOrientationEst)) : " ";

            // check if angular orientation is the end of the string.
            if (trimmedStr.EndsWith(angularOrientationEst))
            {
                //end of the string and no more to look for the meastype
                return $"{bearingNumber}{angularOrientationEst}{measurementType} {machineSide}";
            }


            if (trimmedStr.Length == bearingNumberIndex+2)
            {
                // the char is at the end  
                measurementType = trimmedStr[trimmedStr.Length-1].ToString();
            }
            else if(measurementType == string.Empty)
            {
                // check if not end of string before moving to next
                measurementType = trimmedStr[bearingNumberIndex+2].ToString(); 
            }
            measurementType = MeasurementType.EstimationCodes.Exists(x => x.Equals(measurementType)) 
                ? MeasurementType.EstimationCodes.Single(x => x.Equals(measurementType)) 
                : " ";
            var substring =
                $"{bearingNumber}{angularOrientationEst}{measurementType}";

            var newName = $"{substring} {machineSide}";
            
            return newName;

        }

        private static Match GetAvailableNumberInString(string trimmedStr)
        {
            var numerRegex = new Regex(@"(\s\d+)|(\d+)");
            var match = numerRegex.Match(trimmedStr);
            return match;
        }

        /// <summary>
        /// Revemo all spaces and zeros
        /// </summary>
        /// <returns></returns>
        private string TrimName(string str)
        {
            string trimmedStr = Regex.Replace( str, @"\s", "" );
            return trimmedStr.Replace("0", "");
        }

        private bool IsEndOfString(string str)
        {
            //str.las
            return false;
        }

        private string GetMeasurementType(string str, out string trimmedString)
        {
            trimmedString = str;

            var pattern = new Regex($"(ACC)|(ENV)|(VEL)|(E1)|(E2)|(E3)|(E4)");
            var match = pattern.Match(str);
            
            var index = match.Groups[0].Index;
            trimmedString = str.Remove(index, match.Length);
            switch (match.Value)
            {
                case "ACC":
                    return MeasurementType.Codes[1].Text;
                case  "ENV":
                    return MeasurementType.Codes[5].Text;
                case "VEL":
                    return MeasurementType.Codes[2].Text;
                
                default:
                    return match.Value;
            }
        }

        private string GetMachineSide(string bearingNumber, string code)
        {
            if (code != "") return code;
            switch (bearingNumber)
            {
                case "01":
                case "04":
                    return MachineSide.EstimationCodes[0];
                case "02":
                case "03":
                    return MachineSide.EstimationCodes[1];                
                default:
                    return code;
            }
        }


        private static string GetCode(string str, List<string> codes)
        {
            
            foreach (var code in codes)
            {
                var pattern = new Regex($"(\\s{code}\\s)|({code})|({code}\\s)|\\s({code})$");
                var match = pattern.Match(str);

                if (!match.Success) continue;
                return match.Value;
            }

            return "";
        }


        public IEnumerable<TreeElem> GetWorkSpacePoints(uint workSpaceId, uint tblSetId)
        {
            return _connection.Query<TreeElem>($"SELECT DISTINCT * FROM TREEELEM WHERE CONTAINERTYPE=4 AND TBLSETID={tblSetId} and Parentid!=2147000000 AND HierarchyId={workSpaceId} AND ELEMENTENABLE = 1");
        }
        public void UpdatePoint(EstimateName elem)
        {
            _connection.Execute($"UPDATE TREEELEM SET NAME ='{elem.NewName}' WHERE NAME='{elem.OldName}' AND HIERARCHYID={elem.HierarchyId} AND TREEELEMID={elem.PointId}");

            if (elem.ReferenceId > 0)
            {
                _connection.Execute($"UPDATE TREEELEM SET NAME ='{elem.NewName}' WHERE NAME='{elem.OldName}' AND HIERARCHYID={elem.HierarchyId} AND TREEELEMID={elem.ReferenceId}");
            }
        }
    }
}
