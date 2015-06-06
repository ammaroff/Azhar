using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelCode
{
    public class Sh_1 : StudentRecord

    {
        public Sh_1(ExcelWorkbook excel)
              : base(excel)
        { }

        public override void Set(int CellAddress, string subjectState, bool IsFromLastYear, int? HelpDegOnSubj, object[] degrees)
        {
            if (CellAddress == 1)//القرآن الكريم
            {
                base.Set(13, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 165)//تفسير آيات الأحكام
            {
                base.Set(7, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 166)//علوم الحديث
            {
                base.Set(8, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 167)//اللغة العربية
            {
                base.Set(9, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 168)//تاريخ التشريع
            {
                base.Set(10, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 169)//اللغةالأوربية
            {
                base.Set(11, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 170)//الفقه المقارن
            {
                base.Set(14, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 171)//توحيد
            {
                base.Set(15, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 172)//أصول الفقه
            {
                base.Set(16, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 173)//قضايا فقهية مقارنة
            {
                base.Set(17, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 174)//قاعة بحث أصول فقه
            {
                base.Set(18, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 175)//الفقه
            {
                base.Set(19, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 176)//قاعة بحث فقه
            {
                base.Set(20, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            
        }

    }
}