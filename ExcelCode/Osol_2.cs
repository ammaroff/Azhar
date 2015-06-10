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
        public Osol_2()
            :base("Osol_2.xlsx","RasdOs2.xlsx")
        {

            
        }
        public override int ClassId
        {
            get
            {
                return 2;
            }

           
        }
        public override string SubjYName
        {
            get
            {
                return "س2";
            }


        }


        //        public override void Set(int CellAddress, dynamic row)
        //        {
        //            if (CellAddress == 2)//القرآن الكريم
        //            {
        //                base.Set(13, row);
        //                return;
        //            }
        //            if (CellAddress == 20)//علوم قرآن
        //            {
        //                base.Set(7, row);
        //                return;
        //            }
        //            if (CellAddress == 21)//خطابة
        //            {
        //                base.Set(8, row);
        //                return;
        //            }
        //            if (CellAddress == 22)//منطق قديم
        //            {
        //                base.Set(9, row);
        //                return;
        //            }
        //            if (CellAddress == 23)//فلسفة عامة
        //            {
        //                base.Set(10, row);
        //                return;
        //            }
        //            if (CellAddress == 24)//الفقه
        //            {
        //                base.Set(11, row);
        //                return;
        //            }
        //            if (CellAddress == 25)//شبهات حول السنة
        //            {
        //                base.Set(12, row);
        //                return;
        //            }
        //            if (CellAddress == 26)//توحيد
        //            {
        //                base.Set(14, row);
        //                return;
        //            }
        //            if (CellAddress == 27)//تفسير
        //            {
        //                base.Set(15, row);
        //                return;
        //            }
        //            if (CellAddress == 28)//لغة عربية
        //            {
        //                base.Set(16, row);
        //                return;
        //            }
        //            if (CellAddress == 29)//حديث
        //            {
        //                base.Set(17, row);
        //                return;
        //            }
        //            if (CellAddress == 30)//نظم إسلامية
        //            {
        //                base.Set(18, row);
        //                return;
        //            }
        //            if (CellAddress == 31)//أخلاق
        //            {
        //                base.Set(19, row);
        //                return;
        //            }
        //            if (CellAddress == 32)//تيارات فكرية
        //            {
        //                base.Set(20, row);
        //                return;
        //            }
        //            if (CellAddress == 33)//علوم الحديث
        //            {
        //                base.Set(21, row);
        //                return;
        //            }
        //            if (CellAddress == 34)//شبهات حول القرآن
        //            {
        //                base.Set(22, row);
        //                return;
        //            }
        //			 if (CellAddress == 35)//اللغة الأوربية
        //            {
        //                base.Set(23, row);
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