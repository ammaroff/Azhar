using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelCode
{
    public class Sh_2 : StudentRecord
       
    {
       
        public Sh_2(ExcelWorkbook excel)
              : base(excel)
        { }
        public override void Set(int CellAddress, string subjectState, bool IsFromLastYear, int? HelpDegOnSubj, object[] degrees)
        {
            if (CellAddress == 2)//القرآن الكريم
            {
                base.Set(13, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 177)//تفسير آيات الأحكام
            {
                base.Set(7, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 178)//توحيد
            {
                base.Set(8, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 179)//اللغة العربية
            {
                base.Set(9, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 180)//قضايا فقهية معاصرة
            {
                base.Set(10, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 181)//أحوال شخصية
            {
                base.Set(14, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 182)//الفقه
            {
                base.Set(15, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 183)//أحاديث الأحكام
            {
                base.Set(16, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 184)//اللغة الأوربية
            {
                base.Set(17, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 185)//أصول الفقه
            {
                base.Set(18, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 186)//الفقه المقارن
            {
                base.Set(19, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 187)//قاعة بحث فقه مقارن
            {
                base.Set(20, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 188)//قاعة بحث فقه
            {
                base.Set(21, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }

        }

    }
}