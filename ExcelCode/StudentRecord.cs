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
    }
    public class StudentRecord
    {
        public string SubjectName { get; set; }
        public string Irregular { get; set; }
        public int SheetNumber { get; set; }
        public StudentRecord(ExcelWorkbook excel)
        {

            SheetNumber = 0;
            Sheet = excel.Worksheets.ElementAtOrDefault(SheetNumber);
        }
        public StudentRecord(string subjectName, string seatNo, string studentName, string irregular, string recStatus, int secretNo, string stdState)
        {
            SubjectName = subjectName; SeatNo = seatNo; StudentName = studentName; Irregular = irregular; RecordStatus = recStatus; SecretNo = secretNo; StdState = stdState;

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
        public const int inc = 8;
        public const int start = 5;
        public int current = start;
        public int lastYearIndex = 25;


        public virtual void SetStudet(int index)
        {
            current += (index * inc);
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

        public virtual void Set(int s_index, string subjectState, bool IsFromLastYear, int? HelpDegOnSubj, int OrignalSubjId, object[] degrees)
        {


            for (int i = current; i < degrees.Length + current; i++)
            {
                int cellIndex = i - current;

                int val = -1;
                if (int.TryParse((string)degrees[i - current], out val))
                {

                    if (new int[] { 1, 2, 3, 4 }.Contains(OrignalSubjId))
                    {
                        if (ApplyQuran(s_index, subjectState, i, cellIndex, val)) continue; // درجات القران بالجبر
                        if (ApplyDegrees(s_index, subjectState,IsFromLastYear,HelpDegOnSubj, i, cellIndex, val,degrees[i+1-current])) continue;
                    }
                    else
                    {
                        if (cellIndex == 1 || cellIndex == 2) continue; //لا تضع أي درجة في خانات الجبر

                        Sheet.Cells[i, s_index].Value = val;


                    }


                }
                else
                {
                    if (cellIndex == 1 || cellIndex == 2) continue; //لا تضع أي درجة في خانات الجبر
                    string value = (string)degrees[i - current];
                    Sheet.Cells[i, s_index].Value = value;
                    ApplyGardes(s_index, subjectState, HelpDegOnSubj, degrees, i, cellIndex, value); // ضع التقدير

                }

                //string styleName = styles[i - current];
                //if (!string.IsNullOrWhiteSpace(styleName))
                //{
                //    Sheet.Cells[i, s_index].StyleName = styleName;
                //}








            }
        }

        private void ApplyGardes(int s_index, string subjectState, int? HelpDegOnSubj, object[] degrees, int i, int cellIndex, string value)
        {
            if (cellIndex == 6) //cellIndex 7  التقدير
            {
                if (subjectState == "Fail")
                    Sheet.Cells[i, s_index].WithStyle("FailSubj").Value = (string)degrees[i + 1 - current];
                else
                {
                    if (!HelpDegOnSubj.HasValue || HelpDegOnSubj.Value < 0)
                    {
                        Sheet.Cells[i, s_index].WithStyle("DeNewYearGrade").Value = value;
                    }
                }

            }
            if (cellIndex == 7) //cellIndex  8 التقدير
            {

                if (HelpDegOnSubj.HasValue && HelpDegOnSubj.Value > 0)
                {
                    Sheet.Cells[i, s_index].WithStyle("HelpedSubjGrade").Value = value;
                }


            }
        }

        private bool ApplyDegrees(int s_index, string subjectState,bool IsFromLastYear,int? HelpDegOnSubj, int i, int cellIndex, int total,object lastTotal)
        {
            if (subjectState != "Fail")
            {

                if (HelpDegOnSubj.HasValue && HelpDegOnSubj.Value > 0 && IsFromLastYear)
                {

                    Sheet.Cells[i, s_index].WithStyle("DeNewYearDgree(N)Old").Value = total;
                    return true;

                }
                if (!IsFromLastYear)
                {
                    Sheet.Cells[i, s_index].WithStyle("DeNewYearDgree(N)").Value = total;
                    return true ;
                }

            }
            else
            {
                Sheet.Cells[i, s_index].WithStyle("FailSubj").Value = lastTotal;

            }
            return false;
        }

        private bool ApplyQuran(int s_index, string subjectState, int i, int cellIndex, int val)
        {
            if (new string[] { "Help", "Auto" }.Contains(subjectState))
            {
                if (cellIndex == 0 || cellIndex == 2)
                {
                    if (val < 25)
                    {
                        Sheet.Cells[i, s_index].Value = 25;
                        return true;
                    }


                    else
                    if (val != 0)
                    {
                        Sheet.Cells[i, s_index].Value = val;
                        return true;
                    }
                }
                if (cellIndex == 1 || cellIndex == 3)
                {
                    if (val < 25)
                    {
                        Sheet.Cells[i, s_index].WithStyle("HelpedSubjDegree").Value = val;
                        return true;
                    }
                }
                

            }
            return false;

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



        public virtual void Set(int CellAddress, string subjectState, bool IsFromLastYear, int? HelpDegOnSubj, object[] degrees)
        {
            throw new Exception("لازم تنفذها");

        }
    }
}
