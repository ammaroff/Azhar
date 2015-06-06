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


        public override void Set(int CellAddress, object[] degrees,string [] styles)
        {
            if (CellAddress == 2)//القرآن الكريم
            {
                base.Set(13, degrees, styles);
                return;
            }
            if (CellAddress == 20)//علوم قرآن
            {
                base.Set(7, degrees, styles);
                return;
            }
            if (CellAddress == 21)//خطابة
            {
                base.Set(8, degrees, styles);
                return;
            }
            if (CellAddress == 22)//منطق قديم
            {
                base.Set(9, degrees, styles);
                return;
            }
            if (CellAddress == 23)//فلسفة عامة
            {
                base.Set(10, degrees, styles);
                return;
            }
            if (CellAddress == 24)//الفقه
            {
                base.Set(11, degrees, styles);
                return;
            }
            if (CellAddress == 25)//شبهات حول السنة
            {
                base.Set(12, degrees, styles);
                return;
            }
            if (CellAddress == 26)//توحيد
            {
                base.Set(14, degrees, styles);
                return;
            }
            if (CellAddress == 27)//تفسير
            {
                base.Set(15, degrees, styles);
                return;
            }
            if (CellAddress == 28)//لغة عربية
            {
                base.Set(16, degrees, styles);
                return;
            }
            if (CellAddress == 29)//حديث
            {
                base.Set(17, degrees, styles);
                return;
            }
            if (CellAddress == 30)//نظم إسلامية
            {
                base.Set(18, degrees, styles);
                return;
            }
            if (CellAddress == 31)//أخلاق
            {
                base.Set(19, degrees, styles);
                return;
            }
            if (CellAddress == 32)//تيارات فكرية
            {
                base.Set(20, degrees, styles);
                return;
            }
            if (CellAddress == 33)//علوم الحديث
            {
                base.Set(21, degrees, styles);
                return;
            }
            if (CellAddress == 34)//شبهات حول القرآن
            {
                base.Set(22, degrees, styles);
                return;
            }
			 if (CellAddress == 35)//اللغة الأوربية
            {
                base.Set(23, degrees, styles);
                return;
            }
/*            if (CellAddress < 11 && CellAddress > 4)
            {
                base.Set(CellAddress + 2, degrees, styles);
                return;
            }
            if (CellAddress < 20 && CellAddress > 10)
            {
                base.Set(CellAddress + 3, degrees, styles);
                return;
            }*/

             }

    }
}