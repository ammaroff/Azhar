using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelCode
{
    public class Osol_1 : StudentRecord
    {
        public Osol_1(ExcelWorkbook excel)
             : base(excel)
        {

         
        }



        public override void Set(int CellAddress,string subjectState, bool IsFromLastYear, int? HelpDegOnSubj, object[] degrees)
        {
            if (CellAddress == 1)//قران
            {
                base.Set(13, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 5)//لغة عربية
            {
                base.Set(7, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 6)//نظم إسلامية
            {
                base.Set(8, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 7)//علوم القرآن
            {
                base.Set(9, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 8)//منطق قديم
            {
                base.Set(10, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 9)//فقه
            {
                base.Set(11, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 10)//تاريخ السنة
            {
                base.Set(12, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 11)//تفسير تحليلي
            {
                base.Set(14, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 12)//حديث تحليلي
            {
                base.Set(15, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 13)//توحيد
            {
                base.Set(16, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 14)//أصول الدعوة
            {
                base.Set(17, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 15)//تصوف
            {
                base.Set(18, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 16)//ملل ونحل
            {
                base.Set(19, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 17)//قصص القرآن
            {
                base.Set(20, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 18)//علوم الحديث
            {
                base.Set(21, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 19)//اللغة الأوربية
            {
                base.Set(22, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
/*            if (CellAddress < 11 && CellAddress > 4)
            {
                base.Set(CellAddress + 2, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress < 20 && CellAddress > 10)
            {
                base.Set(CellAddress + 3, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }*/

             }

    }
}