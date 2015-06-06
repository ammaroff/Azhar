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



        public override void Set(int CellAddress, object[] degrees,string [] styles)
        {
            if (CellAddress == 1)//قران
            {
                base.Set(13, degrees, styles);
                return;
            }
            if (CellAddress == 5)//لغة عربية
            {
                base.Set(7, degrees, styles);
                return;
            }
            if (CellAddress == 6)//نظم إسلامية
            {
                base.Set(8, degrees, styles);
                return;
            }
            if (CellAddress == 7)//علوم القرآن
            {
                base.Set(9, degrees, styles);
                return;
            }
            if (CellAddress == 8)//منطق قديم
            {
                base.Set(10, degrees, styles);
                return;
            }
            if (CellAddress == 9)//فقه
            {
                base.Set(11, degrees, styles);
                return;
            }
            if (CellAddress == 10)//تاريخ السنة
            {
                base.Set(12, degrees, styles);
                return;
            }
            if (CellAddress == 11)//تفسير تحليلي
            {
                base.Set(14, degrees, styles);
                return;
            }
            if (CellAddress == 12)//حديث تحليلي
            {
                base.Set(15, degrees, styles);
                return;
            }
            if (CellAddress == 13)//توحيد
            {
                base.Set(16, degrees, styles);
                return;
            }
            if (CellAddress == 14)//أصول الدعوة
            {
                base.Set(17, degrees, styles);
                return;
            }
            if (CellAddress == 15)//تصوف
            {
                base.Set(18, degrees, styles);
                return;
            }
            if (CellAddress == 16)//ملل ونحل
            {
                base.Set(19, degrees, styles);
                return;
            }
            if (CellAddress == 17)//قصص القرآن
            {
                base.Set(20, degrees, styles);
                return;
            }
            if (CellAddress == 18)//علوم الحديث
            {
                base.Set(21, degrees, styles);
                return;
            }
            if (CellAddress == 19)//اللغة الأوربية
            {
                base.Set(22, degrees, styles);
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