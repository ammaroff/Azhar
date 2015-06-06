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
    class Program
    {
        static IEnumerable<dynamic> connect(int ClassId)
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);

            string cmd = "GetAllStdDegByClass";
            return Dapper.SqlMapper.Query(con, cmd, new { ClassId = ClassId }, commandType: CommandType.StoredProcedure);


        }
        static void Main(string[] args)
        {




            Fill_Student_Data(1, typeof(Osol_1));

            Fill_Student_Data(2, typeof(Osol_2));


        }

        private static void Fill_Student_Data(int ClassId, Type type)
        {

            FileInfo newFile = new FileInfo(type.Name + ".xlsx");
            
            Console.WriteLine("Openning " + newFile.Name + " file template");
            using (ExcelPackage pkg = new ExcelPackage(newFile))
            {

                Console.WriteLine("create excel file");
                var data = connect(ClassId);
                int i = 0;
                var groups = data.GroupBy(rows => rows.Num);
                Console.WriteLine("getting data from database for classid {0} and total students  : {1}", ClassId, groups.Count());
                var record = (StudentRecord)Activator.CreateInstance(type,pkg.Workbook);
           
                groups.ToList().ForEach(rows =>
                {


                    string currentStudent = rows.Key.ToString();
                    Console.WriteLine("dump data for student id:{0}", currentStudent);
                    record.SeatNo = currentStudent;
                    record.StudentName = rows.First().StdName;
                    record.Irregular = rows.First().IsIrregular;
                    record.RecordStatus = rows.First().StdType;
                    record.SecretNo = rows.First().SecrtNum;
                    record.StdState = rows.First().StdState;
                    record.SetStudet(i++);



                    string total = rows.First().TotalDeg;
                    string oldTotal = rows.First().TotalBefore;
                    int Isfinal = Convert.ToInt32(rows.First().IsFinal);
                    string StdGrade = rows.First().StdGrade;

                    foreach (var row in rows)
                    {
                        if (row.SubjYearId != ClassId)
                        {
                            //تخلفات 
                            record.SetLastYearSubject(row.SubjName, row.SubjYName, new object[] {
                            row.OralDeg, null, row.WriringDeg, null,
                           row.LastTotal,
                            row.Total==row.LastTotal? null:row.Total,
                           row.LastGrade,
                           row.Grade==row.LastGrade?null:row.Grade
                            });

                        }
                        else
                        {
                            record.Set(row.SubjId, new object[] {
                            row.OralDeg, null, row.WriringDeg, null,
                           row.LastTotal,
                            row.Total==row.LastTotal? null:row.Total,
                           row.LastGrade,
                           row.Grade==row.LastGrade?null:row.Grade
                            },
                           //styles here 
                           new string[] {
                                row.subjectState=="Fail"?"FailSubj":null
                            ,null
                            ,null
                            ,null
                            ,null
                            , row.Total==row.LastTotal?null:"HelpedSubjDegree"
                            , null
                            ,row.LastGrade==row.Grade?null:"HelpedSubjGrade"}
                           );
                        }


                    }
                    record.SetTotal(Isfinal, total, oldTotal);
                    record.SetGrade(Isfinal, StdGrade, StdGrade);
                   
                });
                pkg.Save();
            }
        }
    }
}
