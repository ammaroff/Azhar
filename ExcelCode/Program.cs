using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.IO;
using System.Diagnostics;
using Dapper;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace ExcelCode
{
    public class Program
    {
        public static IEnumerable<dynamic> connect(int ClassId)
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);

            string cmd = "GetAllStdDegByClass";
            return Dapper.SqlMapper.Query(con, cmd, new { ClassId = ClassId }, commandType: CommandType.StoredProcedure);


        }
        static void Main(string[] args)
        {




            Fill_Student_Data(typeof(Osol_1));
            Fill_Student_Data(typeof(Osol_2));
            Fill_Student_Data(typeof(Sh_1));
            Fill_Student_Data(typeof(Sh_2));


        }

        private static void Fill_Student_Data(Type type)
        {





            var record = (StudentRecord)Activator.CreateInstance(type);
            Console.WriteLine("create excel file");
            var data = connect(record.ClassId);
            int i = 0;
            var groups = data.GroupBy(rows => rows.Num);
            Console.WriteLine("getting data from database for classid {0} and total students  : {1}", record.ClassId, groups.Count());


            groups.ToList().ForEach(rows =>
            {


                string currentStudent = rows.Key.ToString();
                Console.WriteLine("dump data for student id:{0}", currentStudent);
                record.SeatNo = currentStudent;
                record.StudentName = rows.First().StdName;
                record.Irregular = rows.First().IsIrregular;
                record.RecordStatus = rows.First().Des;
                record.SecretNo = rows.First().SecrtNum;
                record.StdState = rows.First().StdState;
                record.TotalDeg = rows.First().TotalDeg;
                record.Mazhab = rows.First().Mazhab;
                record.SetStudet(i++);
                record.SetNewRasdStudent();



                string total = rows.First().TotalDeg;
                string oldTotal = rows.First().TotalBefore;
                string Mazhab = rows.First().Mazhab;
                int Isfinal = Convert.ToInt32(rows.First().IsFinal);
                string StdGrade = rows.First().StdGrade;
                string oldStdGrade = rows.First().TotalGradeBefore;


                record.SetGroup(rows);
                record.SetRasdGroup(rows);
                record.SetTotal(Isfinal, total, oldTotal);
                record.SetGrade(Isfinal, StdGrade, oldStdGrade);



            });
            record.Save();


        }
    }
}
