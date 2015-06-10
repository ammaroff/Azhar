using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelCode;
using OfficeOpenXml;

namespace GradesOnly
{
    class Program
    {
        
        static void Main(string[] args)
        {
           
        }
        private static void Fill_Student_Data(Type type)
        {

            FileInfo newFile = new FileInfo(type.Name + ".xlsx");

            Console.WriteLine("Openning " + newFile.Name + " file template");
            using (ExcelPackage pkg = new ExcelPackage(newFile))
            {
                var record = (StudentRecord)Activator.CreateInstance(type, pkg.Workbook);
                Console.WriteLine("create excel file");
                var data = ExcelCode.Program.connect(record.ClassId);
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
                    record.SetStudet(i++);



                    string total = rows.First().TotalDeg;
                    string oldTotal = rows.First().TotalBefore;
                    int Isfinal = Convert.ToInt32(rows.First().IsFinal);
                    string StdGrade = rows.First().StdGrade;
                    string oldStdGrade = rows.First().TotalGradeBefore;


                    record.SetGroup(rows);
                    record.SetTotal(Isfinal, total, oldTotal);
                    record.SetGrade(Isfinal, StdGrade, oldStdGrade);



                });

                pkg.Save();
            }
        }
    }
}
