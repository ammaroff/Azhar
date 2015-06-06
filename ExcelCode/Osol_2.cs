using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelCode
{
    public class Osol_2 : StudentRecord
    {
        public Osol_2(ExcelWorkbook excel)
            :base(excel)
        {

            
        }


        public override void Set(int CellAddress,string subjectState, bool IsFromLastYear, int? HelpDegOnSubj, object[] degrees)
        {
            if (CellAddress == 2)//القرآن الكريم
            {
                base.Set(13, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 20)//علوم قرآن
            {
                base.Set(7, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 21)//خطابة
            {
                base.Set(8, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 22)//منطق قديم
            {
                base.Set(9, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 23)//فلسفة عامة
            {
                base.Set(10, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 24)//الفقه
            {
                base.Set(11, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 25)//شبهات حول السنة
            {
                base.Set(12, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 26)//توحيد
            {
                base.Set(14, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 27)//تفسير
            {
                base.Set(15, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 28)//لغة عربية
            {
                base.Set(16, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 29)//حديث
            {
                base.Set(17, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 30)//نظم إسلامية
            {
                base.Set(18, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 31)//أخلاق
            {
                base.Set(19, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 32)//تيارات فكرية
            {
                base.Set(20, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 33)//علوم الحديث
            {
                base.Set(21, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
            if (CellAddress == 34)//شبهات حول القرآن
            {
                base.Set(22, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
                return;
            }
			 if (CellAddress == 35)//اللغة الأوربية
            {
                base.Set(23, subjectState,IsFromLastYear,HelpDegOnSubj,CellAddress, degrees);
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