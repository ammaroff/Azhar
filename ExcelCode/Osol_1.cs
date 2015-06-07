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
        public override string SubjYName
        {
            get
            {
                return "س1";
            }

           
        }



        //        public override void Set(int CellAddress,dynamic row)
        //        {
        //            if (CellAddress == 1)//قران
        //            {
        //                base.Set(13, row);
        //                return;
        //            }
        //            if (CellAddress == 5)//لغة عربية
        //            {
        //                base.Set(7, row);
        //                return;
        //            }
        //            if (CellAddress == 6)//نظم إسلامية
        //            {
        //                base.Set(8, row);
        //                return;
        //            }
        //            if (CellAddress == 7)//علوم القرآن
        //            {
        //                base.Set(9, row);
        //                return;
        //            }
        //            if (CellAddress == 8)//منطق قديم
        //            {
        //                base.Set(10, row);
        //                return;
        //            }
        //            if (CellAddress == 9)//فقه
        //            {
        //                base.Set(11, row);
        //                return;
        //            }
        //            if (CellAddress == 10)//تاريخ السنة
        //            {
        //                base.Set(12, row);
        //                return;
        //            }
        //            if (CellAddress == 11)//تفسير تحليلي
        //            {
        //                base.Set(14, row);
        //                return;
        //            }
        //            if (CellAddress == 12)//حديث تحليلي
        //            {
        //                base.Set(15, row);
        //                return;
        //            }
        //            if (CellAddress == 13)//توحيد
        //            {
        //                base.Set(16, row);
        //                return;
        //            }
        //            if (CellAddress == 14)//أصول الدعوة
        //            {
        //                base.Set(17, row);
        //                return;
        //            }
        //            if (CellAddress == 15)//تصوف
        //            {
        //                base.Set(18, row);
        //                return;
        //            }
        //            if (CellAddress == 16)//ملل ونحل
        //            {
        //                base.Set(19, row);
        //                return;
        //            }
        //            if (CellAddress == 17)//قصص القرآن
        //            {
        //                base.Set(20, row);
        //                return;
        //            }
        //            if (CellAddress == 18)//علوم الحديث
        //            {
        //                base.Set(21, row);
        //                return;
        //            }
        //            if (CellAddress == 19)//اللغة الأوربية
        //            {
        //                base.Set(22, row);
        //                return;
        //            }
        ///*            if (CellAddress < 11 && CellAddress > 4)
        //            {
        //                base.Set(CellAddress + 2, row);
        //                return;
        //            }
        //            if (CellAddress < 20 && CellAddress > 10)
        //            {
        //                base.Set(CellAddress + 3, row);
        //                return;
        //            }*/

        //             }

    }
}