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
        public static void AddNote(this ExcelWorksheet sheet, int rowIndex, string note, params object[] p)
        {
            if (note == null) return;

            note = string.Format(note, p);
            rowIndex = (((rowIndex - 5) / 8) * 8) + 5;
            sheet.Cells["AJ" + rowIndex.ToString()].Style.WrapText = true;
            
            string[] specialNotes = {"تحويل الي خارج الكلية",
                                        "فصل طالب",
                                        "عذر مرضي مقبول عن مواد الترم الثاني",
                                        "عذر مرضي مقبول عن مواد الترم الثاني",
                                        "طالب ترك الدراسة"
                                        ,"غائب بدون عذر"};
          
            foreach(string sp in specialNotes)
            {
                if (sheet.Cells["AJ" + rowIndex.ToString()].Text.Trim().Contains(sp))
                {
                    sheet.Cells["AJ" + rowIndex.ToString()].Value = sp;
                    return;
                }
                    

            }
            
           

           

                if (sheet.Cells["AJ" + rowIndex.ToString()].Value == null)
                    sheet.Cells["AJ" + rowIndex.ToString()].Value = "";
                sheet.Cells["AJ" + rowIndex.ToString()].Value += Environment.NewLine + note;
           
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
        public virtual string SubjYName { get { return ""; } }
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
        public virtual int ClassId { get { return 0; } }

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
            if (Isfinal == 1)
            {
                if (total == oldTotal)
                {
                    Sheet.Cells[current, 29].Value = total;
                }

                else
                {
                    Sheet.Cells[current, 29].Value = total;
                    Sheet.Cells[current + 4, 29].Value = oldTotal;
                    Sheet.Cells[current + 4, 29].StyleName = "TotalDegreeHelp";
                }

            }
        }

        public virtual void SetGrade(int Isfinal, string StdGrade, string oldStdGrade)
        {

            if (Isfinal == 1)
            {

                if (StdGrade == oldStdGrade)
                {
                    Sheet.Cells[current, 30].Value = StdGrade;
                }
                else
                {
                    Sheet.Cells[current, 30].Value = StdGrade;
                    Sheet.Cells[current + 4, 30].Value = oldStdGrade;
                    Sheet.Cells[current + 4, 30].StyleName = "TotalDegreeHelp";
                }
            }

        }

        private void Set(dynamic row)
        {


            #region إعدادات
            int subjid = row.SubjId;
            string subjName = row.SubjName;
            var subject = Sheet.Cells["G3:W3"].FirstOrDefault(cell => cell.Text == subjName);
            int i = -1;
            #endregion

            if (row.SubjYName != this.SubjYName || subject == null)
            {
                #region مادة التخلف الأولى
                if (Sheet.Cells["Y" + current.ToString()].Value == null)
                {

                    //اسم المادة
                    Sheet.Cells["Y" + current.ToString()].Value = subjName;
                    Sheet.Cells["Y" + (current + 7).ToString()].Value = row.SubjYName;

                    SetDegrees(subjid, subjName, row, i, 26);

                }
                #endregion
                #region مادة التخلف الثانية
                else
                    if (Sheet.Cells["AA" + current.ToString()].Value == null)
                {
                    //اسم المادة
                    Sheet.Cells["AA" + current.ToString()].Value = subjName;
                    Sheet.Cells["AA" + (current + 7).ToString()].Value = row.SubjYName;

                    SetDegrees(subjid, subjName, row, i, 28);

                }

                #endregion
                return;

            }
            int colIndex = subject.Start.Column;

            SetDegrees(subjid, subjName, row, i, colIndex);

        }

        private static string strDegrees(double degree)
        {
            if (degree == 1)
                return "درجة واحدة";
            if (degree == 2)
                return "درجتين";
            if (degree < 11)
                return string.Format("{0} درجات",degree);
            return string.Format("{0} درجة", degree);

        }
        bool anyHelp = false;

        public virtual void SetGroup(IGrouping<dynamic, dynamic> rows)
        {

            StdState = rows.FirstOrDefault().StdState;
            #region ملاحظات خاصة كالفصل والعذر وغيره
            Sheet.AddNote(current, (string)rows.FirstOrDefault().specialnotes);
            #endregion


            #region الغائب
            if (rows
                .Where(i => !i.IsFromLastYear)
                .All(i => i.Grade == "غ"))
            {
                Sheet.Cells[current, 31].Value = "غائب ";
                Sheet.AddNote(current, "غائب بدون عذر");

            }
            #endregion

            foreach (var row in rows)
            {

                Set(row);
            }

            #region حالة الجبر
            // في حالة الجبر في مادة أو مادتين: جبر بــ (عدد الدرجات) في(اسم المادة) لينجح بتقدير
            //If row.HelpDegOnSubj > 0 ? "جبر بـ " + row.HelpDegOnSubj + " في " + row.SubjName + " " + row.SubjYName + " - "
            //لو كان هناك مادة ثانية وثالثة يضافون بنفس الطريقة
            //ثم تضاف في النهاية كلمة + "لينجح بتقدير"

            foreach (var row in rows.Where(i => i.HelpDegOnSubj > 0 && i.SubjName != "القرآن الكريم" && i.subjectState == "Help"))
            {
                anyHelp = true;
                //الجبر الذي يكتب في الملاحظات هو الجبر الذي فوق درجتين .. يمكن معرفة ذلك من خلال 
                // subjectState = Help
                //وهذا يصلح في كل المواد إلا مادة القرآن فلها حسبة خاصة ذكرتها لك
                // أما الجبر الذي تكون حالة المادة معه
                //subjectState =Auto
                //فهذا لا يكتب في الملاحظات
                anyHelp = true;
                string LastYearSubjName = (SubjYName != row.SubjYName) ? row.SubjYName : "";
                Sheet.AddNote(current, "جبر بـ{0}  في {1} {2}", strDegrees((int)row.HelpDegOnSubj), (string)row.SubjName, LastYearSubjName);

            }


            #endregion
            if (anyHelp && !StdState.Contains("منقول"))
                Sheet.AddNote(current, "لينجح بتقدير");



            #region الراسب

            //الكود التالي لا يعمل .. ولابد من الشرطين هذين حتى يتحقق أن الطالب راسب وليس بمنقول
            
            if (!rows.FirstOrDefault().IsFinal && !string.IsNullOrWhiteSpace(StdState) && StdState.Contains("منقول بماد"))//طالب راسب
            {
                // احسب عدد مواد الرسوب
                int failCount = rows.Count(i => i.subjectState == "Fail");
                int maxStayId = rows.FirstOrDefault().MaxStayId;
                if (new int[] { 2, 6, 11 }.Contains(maxStayId))
                {
                    Sheet.Cells[current, 31].Value = "راسب وينظر في فصله ";
                }

                //    if (failCount == 1)
                //       Sheet.AddNote(current, " راسب في مادة واحدة", failCount);
                //    if (failCount == 2)
                //       Sheet.AddNote(current, " راسب في مادتين ", failCount);
                if (failCount > 2 && failCount < 11)
                    Sheet.AddNote(current, " راسب في {0} مواد ", failCount);
                if (failCount > 10)
                    Sheet.AddNote(current, "راسب في {0} مادة", failCount);




            }
            #endregion

           

            #region منقول بمواد
            /*
            //في حالة النقل بمادة أو مادتين بدون جبر
            //مثال للكود المقترح للمادة
            If row.StdStat == "منقول بمادة" ? "منقول بمادة "+ row.SubjName [that has] row.subjctState == "Fail" +" "+ row.SubjYName
            //الكود المقترح للمادتين
            If row.StdStat == "منقول بمادتين" ? "منقول بمادتين "+ row.SubjName [that has] row.subjctState == "Fail" +" "+ row.SubjYName +" و" row.SubjName [that has] row.subjctState == "Fail" +" "+ row.SubjYName
            */
            if (!string.IsNullOrWhiteSpace(StdState) && StdState.Contains("منقول بماد") && rows.Any(i => i.subjectState == "Fail") )
            {

                var subjectsWithHelpFromFailArray = rows.Where(i => i.subjectState == "Fail").Select(i => (string)i.SubjName + " " + (((string)i.SubjYName) == this.SubjYName ? "" : ((string)i.SubjYName))).ToArray();
                var subjectsWithHelpFromFail = string.Join(" و", subjectsWithHelpFromFailArray);
                Sheet.AddNote(current, (string)rows.FirstOrDefault().StdState + " " + subjectsWithHelpFromFail);

            }

            #endregion

            #region منح الدرجة الأعلى
            ///في حالة المنح في المجموع الكلي .. الكود المقترح
            // If row.HelpDegOnTotalDeg > 0 ? "منح " + row.HelpDegOnTotalDeg + "درجة أو درجات في المجموع الكلي ليتمتع بالتقدير الأعلى"
            if (rows.FirstOrDefault().HelpDegOnTotalDeg > 0)
            {
                Sheet.AddNote(current, "منح {0}  في المجموع الكلي ليتمتع بالتقدير الأعلى", strDegrees((int)rows.FirstOrDefault().HelpDegOnTotalDeg));
            }
            #endregion

            #region منح النجاح
            //في حالة المنح ليصل للحد الأدنى للنجاح
            //لم أفعله بعد
            //IF(row.IsFinal == 1 && row.TotalBefore < row.HalfMaxTotal) ? "منح " + row.HalfMaxTotal - row.TotalBefore + " في المجموع الكلي ليصل للحد الأدنى للنجاح" : null

            var first = rows.FirstOrDefault();
            int? totalBefore = ((string)first.TotalBefore).Parse<int?>();
            if (first.IsFinal && totalBefore.HasValue && totalBefore < first.HalfMaxTotal)
            {
                Sheet.AddNote(current, "منح {0} في المجموع الكلي ليصل للحد الأدنى للنجاح", (int)first.HalfMaxTotal - totalBefore);
            }
            #endregion

            





        }
        private void SetDegrees(int subjId, string subjName, dynamic row, int i, int colIndex)
        {
            #region الخانة الأولى شفوي
            //الخانة الأولى شفوي بدون جبر
            //
            i++;

            if (subjName == "القرآن الكريم" && row.subjectState != "Fail")
            {
                string oralDeg = row.OralDeg;
                if (oralDeg.Parse<float?>().HasValue && oralDeg.Parse<float?>() < 25)
                {
                    Sheet.Cells[current + i, colIndex].Value = 25;
                }
                else
                {
                    //                    if (row.Oral != 0)
                    Sheet.Cells[current + i, colIndex].Value = row.OralDeg;
                }

            }

            else //باقي المواد
                if (row.Oral != 0)
                Sheet.Cells[current + i, colIndex].Value = row.OralDeg;

            ////////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion


            #region الخانة الثانية .. شفوي جبر القرآن
            //الخانة الثانية
            i++;
            if (subjName == "القرآن الكريم" && row.subjectState != "Fail")
            {
                string oralDeg = row.OralDeg;
                if (oralDeg.Parse<float?>().HasValue && oralDeg.Parse<float?>() < 25)
                {
                    Sheet.Cells[current + i, colIndex].WithStyle("HelpedSubjDegree").Value = row.OralDeg;
                }
                if (oralDeg.Parse<float?>().HasValue && oralDeg.Parse<float?>() < 24 && oralDeg.Parse<float?>() > 18 && row.subjectState == "Help")
                {
                    anyHelp = true;
                    double help = 25 - oralDeg.Parse<double>();
                    string LastYearSubjName = (SubjYName != row.SubjYName) ? row.SubjYName : "";
                    Sheet.AddNote(current, "جبر بـ{0}  في شفوي القران الكريم {1}", strDegrees( help), LastYearSubjName);
                }



            }

            //باقي المواد
            //لا تفعل شيئا

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion


            #region الخانة الثالثة تحريري
            //الخانة الثالثة تحريري بدون جبر
            //
            i++;
            string writingDeg = row.WriringDeg;
            if (subjName == "القرآن الكريم" && row.subjectState != "Fail")
            {

                if (writingDeg.Parse<float?>().HasValue && writingDeg.Parse<float?>() < 25)
                {
                    Sheet.Cells[current + i, colIndex].Value = 25;
                }
                else
                {
                    Sheet.Cells[current + i, colIndex].Value = writingDeg;
                }

            }

            else //باقي المواد
                Sheet.Cells[current + i, colIndex].Value = writingDeg;

            ////////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion


            #region الخانة الرابعة تحريري جبر القرآن
            //الخانة الرابعة تحريري الجبر
            //
            i++;
            if (subjName == "القرآن الكريم" && row.subjectState != "Fail")
            {


                if (writingDeg.Parse<float?>().HasValue && writingDeg.Parse<float?>() < 25)
                {
                    Sheet.Cells[current + i, colIndex].WithStyle("HelpedSubjDegree").Value = writingDeg;
                }
                if (writingDeg.Parse<float?>().HasValue && writingDeg.Parse<float?>() < 24 && writingDeg.Parse<float?>() > 18)
                {
                    anyHelp = true;
                    float help = 25 - writingDeg.Parse<float>();
                    string LastYearSubjName = (SubjYName != row.SubjYName) ? row.SubjYName : "";
                    Sheet.AddNote(current, "جبر بـ{0}  في تحريري القران الكريم {1}", strDegrees( help), LastYearSubjName);
                }


            }

            //باقي المواد
            //لا تفعل شيئا

            ////////////////////////////////////////////////////////////////////////////////////////////////////
            #endregion

            #region الخانة الخامسة .. المجموع النهائي

            //
            i++;


            if (subjName == "القرآن الكريم" && row.subjectState == "Fail")


            { }



            else
            {


                if (row.IsFromLastYear)  // مادة من العام الماضي
                {
                    if (row.HelpDegOnSubj < 0) // درجة رأفة بالتقص
                        Sheet.Cells[current + i, colIndex].WithStyle("DeNewYearDgree(N)Old").Value = row.Total;
                    else // ليس له درجة رافة بالتقص
                    {
                        Sheet.Cells[current + i, colIndex].WithStyle("DegreeLastYear").Value = row.LastTotal;

                    }
                }
                else // مادة جديدة وليست من العام الماضي
                {
                    if (row.HelpDegOnSubj < 0) // له درجة رأفة بالنقص
                    {
                        Sheet.Cells[current + i, colIndex].WithStyle("DeNewYearDgree(N)").Value = row.Total;
                    }
                    else
                        if (row.subjectState == "Fail")
                    {
                        Sheet.Cells[current + i, colIndex].WithStyle("FailSubj").Value = row.LastTotal;
                    }
                    else// إرغامه على شطب مجموع درجة المادة الباقية مع الطالب من العام الماضي وحصل فيها على أعلى من 64
                            if (row.Remaning != null && ((object)(row.Total)).Parse<float?>() > 64)
                    {

                        Sheet.Cells[current + i, colIndex].WithStyle("DeNewYearDgree(N)").Value = row.Total;

                    }
                    else
                        Sheet.Cells[current + i, colIndex].Value = row.LastTotal;

                }


            }



            ////////////////////////////////////////////////////////////////////////////////////////////////////


            #endregion

            #region الخانة السادسة .. المجموع الأصلي

            //
            i++;

            // إرغامه على تنزيل المادة إلى الحد الأعلى للمقبول 64 لأن البرنامج لا يفعلها في بعض الحالات
            if (row.Remaning != null && ((object)(row.Total)).Parse<float?>() > 64)
            {

                Sheet.Cells[current + i, colIndex].Value = 64;

            }
            else

            if (subjName == "القرآن الكريم" && (row.Total == "غ" || ((object)(row.Total)).Parse<float?>() >= 50))
            {

            }


            else
            {


                if (row.HelpDegOnSubj > 0)  // درجة الرأفة بزيادة
                {

                    Sheet.Cells[current + i, colIndex].WithStyle("HelpedSubjDegree").Value = row.Total;
                }
                else // له درجة رأفة بالنقص


                    if (row.HelpDegOnSubj < 0) // له درجة رأفة بالنقص
                {
                    Sheet.Cells[current + i, colIndex].Value = row.LastTotal;
                }


            }



            ////////////////////////////////////////////////////////////////////////////////////////////////////


            #endregion
            #region الخانة السابعة .. التقدير النهائي

            //
            i++;

            if (row.subjectState == "Fail") // إذا كان راسبا في المادة
            {
                Sheet.Cells[current + i, colIndex].WithStyle("FailSubj").Value = row.LastGrade;
            }

            else // إذا كان ناجحا في المادة
            {

                if (row.HelpDegOnSubj < 0)  // وأخذ درجة رأفة بالنقص
                {

                    Sheet.Cells[current + i, colIndex].WithStyle("DeNewYearGrade").Value = row.Grade;
                }
                else // التقدير النهائي في المادة

                    Sheet.Cells[current + i, colIndex].Value = row.LastGrade;
            }


            ////////////////////////////////////////////////////////////////////////////////////////////////////

            #endregion

            #region الخانة الثامنة .. التقدير الأصلي

            //
            i++;
            //إرغامه على رصد تقدير مادة القرآن ضعيف إذا اكتشف أنه مجبور فيها ومجموعها فوق 50 لأن الداتبابيز لا تخفض التقدير
            if (subjName == "القرآن الكريم" && row.subjectState != "Fail" && (Convert.ToInt32(row.OralDeg) < 25 || Convert.ToInt32(row.WriringDeg) < 25))
            {
                Sheet.Cells[current + i, colIndex].WithStyle("HelpedSubjGrade").Value = "ض";
            }
            else // بقية المواد
            {

                if (row.HelpDegOnSubj > 0)  // أخذ درجة رأفة بالزيادة
                {

                    Sheet.Cells[current + i, colIndex].WithStyle("HelpedSubjGrade").Value = row.Grade;
                }
                else // التقدير النهائي في المادة

                    if (row.HelpDegOnSubj < 0) // أخذ درجة رأفة بالنقص
                    Sheet.Cells[current + i, colIndex].Value = row.LastGrade;
            }


            ////////////////////////////////////////////////////////////////////////////////////////////////////

            #endregion

            i++;
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