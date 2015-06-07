using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace ExcelCode
{
    public static class ext
    {
        public static ExcelRange WithStyle(this ExcelRange cells, string styleName)
        {
            cells.StyleName = styleName;
            return cells;
        }
        public static T Parse<T>(this object value)
        {
            try { return (T)System.ComponentModel.TypeDescriptor.GetConverter(typeof(T)).ConvertFrom(value.ToString()); }
            catch { return default(T); }
        }
    }
    public class StudentRecord
    {
        // public string SubjectName { get; set; }
        public string Irregular { get; set; }
        public int SheetNumber { get; set; }
        public virtual string SubjYName { get; }
        public StudentRecord(ExcelWorkbook excel)
        {

            SheetNumber = 0;
            Sheet = excel.Worksheets.ElementAtOrDefault(SheetNumber);
        }
        public StudentRecord(string seatNo, string studentName, string irregular, string recStatus, int secretNo, string stdState)
        {
            SeatNo = seatNo; StudentName = studentName; Irregular = irregular; RecordStatus = recStatus; SecretNo = secretNo; StdState = stdState;

        }
        public StudentRecord(int sheetNumber, string seatNo, string studentName, string recStatus, int secretNo)
        {
            SheetNumber = sheetNumber; SeatNo = seatNo; StudentName = studentName; RecordStatus = recStatus; SecretNo = secretNo;

        }
        public string SeatNo { get; set; }
        public string RecordStatus { get; set; }
        public string StdState { get; set; }
        public int SecretNo { get; set; }
        protected ExcelWorksheet Sheet { get; set; }

        public string StudentName { get; set; }
        public virtual int ClassId { get; }

        public const int inc = 8;
        public const int start = 5;
        public int current = start;
        public int lastYearIndex = 25;


        public virtual void SetStudet(int index)
        {
            current = start + (index * inc);
            Sheet.Cells[current, 1].Value = SeatNo;
            Sheet.Cells[current, 2].Value = StudentName;
            Sheet.Cells[current, 3].Value = Irregular;
            Sheet.Cells[current, 4].Value = RecordStatus;
            Sheet.Cells[current, 5].Value = SecretNo;
            Sheet.Cells[current, 31].Value = StdState;
            lastYearIndex = 25;
        }
        public virtual void SetTotal(int Isfinal, string total, string oldTotal)
        {
            if (Isfinal == 0)
            {
                Sheet.Cells[current, 29].Value = null;
            }
            else
            {
                Sheet.Cells[current, 29].Value = total;
            }
            if (total == oldTotal)
            {
                Sheet.Cells[current + 4, 29].Value = null;
            }
            else
            {
                Sheet.Cells[current + 4, 29].Value = oldTotal;
                Sheet.Cells[current + 4, 29].StyleName = "TotalDegreeHelp";
            }
        }

        public virtual void SetGrade(int Isfinal, string StdGrade, string oldStdGrade)
        {

            if (Isfinal == 0)
            {
                Sheet.Cells[current, 30].Value = null;
            }
            else
            {
                Sheet.Cells[current, 30].Value = StdGrade;
            }
            if (StdGrade == oldStdGrade)
            {
                Sheet.Cells[current + 4, 30].Value = null;
            }
            else
            {
                Sheet.Cells[current + 4, 30].Value = oldStdGrade;
                Sheet.Cells[current + 4, 30].StyleName = "TotalDegreeHelp";
            }
        }

        public virtual void Set(dynamic row)
        {


            #region إعدادات
            int subjid = row.SubjId;
            string subjName = row.SubjName;
            var subject = Sheet.Cells["G3:W3"].FirstOrDefault(cell => cell.Text == subjName);
            int i = -1;

            if (row.SubjYName != this.SubjYName || subject == null)
            {
                #region مادة التخلف الأولى
                if (Sheet.Cells["Y" + current.ToString()].Value == null)
                {

                    //اسم المادة
                    Sheet.Cells["Y" + current.ToString()].Value = subjName;
                    Sheet.Cells["Y" + (current + 7).ToString()].Value = row.SubjYName;

                }
                #endregion
                #region مادة التخلف الثانية
                else
                if (Sheet.Cells["AA" + current.ToString()].Value == null)
                {
                    //اسم المادة
                    Sheet.Cells["AA" + current.ToString()].Value = subjName;
                    Sheet.Cells["AA" + (current + 7).ToString()].Value = row.SubjYName;

                }

                #endregion
                return;

            }
            int colIndex = subject.Start.Column;




            #endregion

            #region ألخانة الأولى شفوي بدون جبر
            //الخانة الأولى شفوي بدون جبر
            //
            i++;
            if (subjName == "القرآن الكريم" && (row.subjectState == "Help" || row.subjectState == "Auto"))
            {
                string oralDeg = row.OralDeg;
                if (oralDeg.Parse<float?>().HasValue && oralDeg.Parse<float?>() < 25)
                {
                    Sheet.Cells[current + i, colIndex].Value = 25;
                }
                else
                {
                    if (row.Oral != 0)
                        Sheet.Cells[current + i, colIndex].Value = row.OralDeg;
                }

            }

            else //باقي المواد
                if (row.Oral != 0)
                Sheet.Cells[current + i, colIndex].Value = row.OralDeg;

            ////////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion


            #region الخانة الثانية
            //الخانة الثانية
            i++;
            if (subjName == "القرآن الكريم" && (row.subjectState == "Help" || row.subjectState == "Auto"))
            {
                string oralDeg = row.OralDeg;
                if (oralDeg.Parse<float?>().HasValue && oralDeg.Parse<float?>() < 25)
                {
                    Sheet.Cells[current + i, colIndex].WithStyle("HelpedSubjDegree").Value = row.OralDeg;
                }


            }

            //باقي المواد
            //لا تفعل شيئا

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion


            #region الخانة الثالثة تحريري بدون جبر
            //الخانة الثالثة تحريري بدون جبر
            //
            i++;
            if (subjName == "القرآن الكريم" && (row.subjectState == "Help" || row.subjectState == "Auto"))
            {
                string writingDeg = row.WriringDeg;
                if (writingDeg.Parse<float?>().HasValue && writingDeg.Parse<float?>() < 25)
                {
                    Sheet.Cells[current + i, colIndex].Value = 25;
                }
                else
                {
                    Sheet.Cells[current + i, colIndex].Value = row.writingDeg;
                }

            }

            else //باقي المواد
                Sheet.Cells[current + i, colIndex].Value = row.writingDeg;

            ////////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion


            #region الخانة الرابعة تحريري الجبر
            //الخانة الرابعة تحريري الجبر
            //
            i++;
            if (subjName == "القرآن الكريم" && (row.subjectState == "Help" || row.subjectState == "Auto"))
            {
                string writingDeg = row.WriringDeg;
                if (writingDeg.Parse<float?>().HasValue && writingDeg.Parse<float?>() < 25)
                {
                    Sheet.Cells[current + i, colIndex].WithStyle("HelpedSubjDegree").Value = writingDeg;
                }


            }

            //باقي المواد
            //لا تفعل شيئا

            ////////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion

            #region الخانة الخامسة 

            //
            i++;


            if (subjName == "القرآن الكريم")
            {
                if (row.subjectState != "Fail")
                {

                    if (row.IsFromLastYear)
                    {
                        if (row.HelpDegOnSubj < 0)
                            Sheet.Cells[current + i, colIndex].WithStyle("DeNewYearDgree(N)Old").Value = row.Total;
                        else
                        {
                            Sheet.Cells[current + i, colIndex].WithStyle("DeNewYearDgree(N)").Value = row.Total;
                        }
                    }
                    else
                    { Sheet.Cells[current + i, colIndex].WithStyle("DegreeLastYear").Value = row.LastTotal; }


                }
                else
                { Sheet.Cells[current + i, colIndex].WithStyle("FailSubj").Value = row.LastTotal; }


            }
            else // باقي المواد
            {
                Sheet.Cells[current + i, colIndex].Value = row.LastTotal;

            }



            ////////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion







        }



        public virtual void SetLastYearSubject(string SubjectName, string Year, object[] degrees)
        {
            if (lastYearIndex > 27) return;
            Sheet.Cells[current, lastYearIndex].Value = SubjectName;
            Sheet.Cells[current + 7, lastYearIndex].Value = Year;
            for (int i = current; i < degrees.Length + current; i++)
            {
                int val = -1;
                if (int.TryParse((string)degrees[i - current], out val))
                {
                    Sheet.Cells[i, lastYearIndex + 1].Value = val;

                }
                else
                {
                    string value = (string)degrees[i - current];
                    Sheet.Cells[i, lastYearIndex + 1].Value = value;

                }

            }

            lastYearIndex += 2;

        }



    }
}
