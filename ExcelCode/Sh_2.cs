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
        public override void Set(int CellAddress, object[] degrees,string [] styles)
        {
            if (CellAddress == 2)//القرآن الكريم
            {
                base.Set(13, degrees, styles);
                return;
            }
            if (CellAddress == 177)//تفسير آيات الأحكام
            {
                base.Set(7, degrees, styles);
                return;
            }
            if (CellAddress == 178)//توحيد
            {
                base.Set(8, degrees, styles);
                return;
            }
            if (CellAddress == 179)//اللغة العربية
            {
                base.Set(9, degrees, styles);
                return;
            }
            if (CellAddress == 180)//قضايا فقهية معاصرة
            {
                base.Set(10, degrees, styles);
                return;
            }
            if (CellAddress == 181)//أحوال شخصية
            {
                base.Set(14, degrees, styles);
                return;
            }
            if (CellAddress == 182)//الفقه
            {
                base.Set(15, degrees, styles);
                return;
            }
            if (CellAddress == 183)//أحاديث الأحكام
            {
                base.Set(16, degrees, styles);
                return;
            }
            if (CellAddress == 184)//اللغة الأوربية
            {
                base.Set(17, degrees, styles);
                return;
            }
            if (CellAddress == 185)//أصول الفقه
            {
                base.Set(18, degrees, styles);
                return;
            }
            if (CellAddress == 186)//الفقه المقارن
            {
                base.Set(19, degrees, styles);
                return;
            }
            if (CellAddress == 187)//قاعة بحث فقه مقارن
            {
                base.Set(20, degrees, styles);
                return;
            }
            if (CellAddress == 188)//قاعة بحث فقه
            {
                base.Set(21, degrees, styles);
                return;
            }

        }

    }
}