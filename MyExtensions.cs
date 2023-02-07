using GubusAdminLTE.MyClasses;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

namespace MyExtensions
{
   
    /// <summary>
    /// Stringler için
    /// </summary>
    public static partial class AllMyExtensions
    {
        public static string mTextLowerAndFirstUpper(this string str)
        {
            str = str.ToLower();
            char[] stra = str.ToCharArray();

            for (int i = 0; i < stra.Length; i++)
            {
                if (i == 0)
                {
                    str = string.Empty;
                    str += stra[i].ToString().ToUpper();
                }
                else
                {
                    str += stra[i].ToString();
                }
            }

            return str;
        }

        public static string mToCamelCase(this string str)
        {
            if (!string.IsNullOrEmpty(str) && str.Length > 1)
                return Char.ToLowerInvariant(str[0]) + str.Substring(1);

            var s = str;

            return s.Replace(" ", "");
        }
    }

    /// <summary>
    /// Objectler için
    /// </summary>
    public static partial class AllMyExtensions
    {
        public static double mToDouble(this object value)
        {
            if (string.IsNullOrEmpty(value.ToString()))
                return 0;

            double d = 0;

            if (double.TryParse(value.ToString().Trim().Replace('.', ','), out double d1))
            {
                d = d1;
            }
            else if (double.TryParse(value.ToString().Trim().Replace(',', '.'), out double d2))
            {
                d = d2;
            }
            else
            {
                d = double.NaN;
            }

            return d;
        }

        public static int mToInt32(this object value)
        {
            if (string.IsNullOrEmpty(value.ToString()))
                return 0;

            return Convert.ToInt32(value);
        }
    }

    /// <summary>
    /// IList için
    /// </summary>
    public static partial class AllMyExtensions
    {
        public static DataTable mToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();

            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, prop.PropertyType);
            }

            object[] values = new object[props.Count];

            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item);
                }
                table.Rows.Add(values);
            }

            return table;
        }
    }

    /// <summary>
    /// DataColumnCollection için
    /// </summary>
    public static partial class AllMyExtensions
    {
        public static void mKolonSil(this DataColumnCollection col, string KolonAdi)
        {
            var c = col[KolonAdi];

            if (c != null)
                col.Remove(KolonAdi);
        }

        public static void mKolonSilVeEkle(this DataColumnCollection col, string KolonAdi)
        {
            mKolonSil(col, KolonAdi);

            col.Add(KolonAdi);
        }
    }

    /// <summary>
    /// DataTable için
    /// </summary>
    public static partial class AllMyExtensions
    {

        public static void mSatirDegerReplace(this DataTable dt, string KolonBasligi, string EskiDeger, string YeniDeger)
        {
            foreach (DataRow row in dt.Rows)
            {
                var s = row[KolonBasligi].ToString().Trim();

                if (s == EskiDeger)
                    row[KolonBasligi] = YeniDeger;
            }

        }

        /// <summary>
        /// Paramslarda verilen ifadeleri içeren satırlar hariç diğerlerini siler.
        /// </summary>
        /// <param name="KolonAdi">Hariçlerin aranacağı kolon adı.</param>
        /// <param name="Haricler">Hariç tutualacak ifadeler.</param>
        public static void mBazilariHaricSatirlariSil(this DataTable dt, string KolonAdi, params string[] Haricler)
        {
            if (dt.Rows.Count > 0)
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow dr = dt.Rows[i];

                    if (Haricler.Contains(dr[KolonAdi].ToString().Trim()))
                        continue;
                    else
                        dr.Delete();
                }
        }

        public static int mDegerinSatirIndexi(this DataTable dt, string KolonAdi, string ArananDeger)
        {
            int t = -1;
            int levelCount = 0;

            try
            {
                if (dt.Rows.Count > 0)
                    for (int i = 0; i < dt.Rows.Count; i++)
                        if (dt.Rows[i][KolonAdi].ToString().Trim() == ArananDeger)
                        {
                            t = i;

                            break;
                        }
            }
            catch
            {
                t = -1;
            }

            return t;
        }

        /// <summary>
        /// Değeri ilgili kolonda en üstten itibaren arar ve ilk bulduğu yeri index olarak verir. Bulunamaması durumunda -1 verir.
        /// </summary>
        public static int mDegerinSatirIndexi(this DataTable dt, int KolonIndex, string ArananDeger)
        {
            int t = -1;
            try
            {
                if (dt.Rows.Count > 0)
                    for (int i = 0; i < dt.Rows.Count; i++)
                        if (dt.Rows[i][KolonIndex].ToString().Trim() == ArananDeger)
                        {
                            t = i;
                            break;
                        }
            }
            catch
            {
                t = -1;
            }

            return t;
        }

        /// <summary>
        /// Dörtlü bosluk, trim tarafıdnan temizlenmeyen boşluğun adıdır. Tek boşluk yerine 4 boşluk eklemek demek <u><strong>değildir</strong></u>.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="KolonAdi"></param>
        public static void mKolondakiBosluklariDortluBoslugaCevir(this DataTable dt, string KolonAdi)
        {
            //Accountdaki boşlukları kısa dörtlü ile değiştir.
            for (int i = 0; i < dt.Rows.Count; i++)
                dt.Rows[i][KolonAdi] = dt.Rows[i][KolonAdi].ToString().Replace(" ", " ");
        }

        /// <summary>
        /// Değeri ilgili kolonda en üstten itibaren arar ve ilk bulduğu yeri index olarak verir.
        /// </summary>
        /// <summary>
        /// Belirtilen isimdeki kolonun tüm satırlarındaki verileri trimler.
        /// </summary>
        public static void mKolondakiMetinleriTrimle(this DataTable dt, string KolonAdi)
        {
            if (dt.Rows.Count > 0)
                foreach (DataRow dr in dt.Rows)
                {
                    var k = dr[KolonAdi].ToString();

                    if (string.IsNullOrEmpty(k))
                        dr[KolonAdi] = "";
                    else
                        dr[KolonAdi] = dr[KolonAdi].ToString().Trim();
                }
        }

        public static void mKolondakiSayilariOndalikliYap(this DataTable dt, string KolonAdi)
        {
            //Accountdaki boşlukları kısa dörtlü ile değiştir.
            for (int i = 0; i < dt.Rows.Count; i++)
                dt.Rows[i][KolonAdi] = MyTools.Convert2MoneyFormat(dt.Rows[i][KolonAdi].ToString());
        }

        public static void mKolonlariSil(this DataTable dt, int HangiKolondanSonrasi, string HaricTutulacakKolonBasligi)
        {//HACK: Geliştirilecek
            if (dt.Columns.Count > 0)
            {
                List<string> list = new List<string>();

                int v = (dt.Columns.Count);
                for (int i = HangiKolondanSonrasi; i < v; i++)
                {
                    string columnName = dt.Columns[i].ColumnName;

                    if (columnName != HaricTutulacakKolonBasligi)
                    {
                        list.Add(columnName);
                    }
                }

                foreach (var l in list)
                {
                    try
                    {
                        dt.Columns.Remove(l);
                    }
                    catch
                    {
                    }
                }
            }
        }


        /// <summary>
        /// Kolonları siler.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="HaricKolon">Silinmeyecek kolon</param>
        /// <param name="HaricKolonlarBaslikIcerigi">Silinemeyecek kolonlarin <strong>içerdiği</strong> ifade</param>
        public static void mKolonlariSil(this DataTable dt, params string[] HaricKolonlarBaslikIcerigi)
        {//HACK: Geliştirilecek
            if (dt.Columns.Count > 0)
            {
                List<string> list = new List<string>();

                int v = (dt.Columns.Count);


                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    string columnName = dt.Columns[c].ColumnName.Trim();

                    if (HaricKolonlarBaslikIcerigi.Any(columnName.Contains))
                        continue;
                    else
                        list.Add(columnName);
                }



                foreach (var l in list)
                {
                    try
                    {
                        dt.Columns.Remove(l);
                    }
                    catch
                    {
                    }
                }
            }
        }


        /// <summary>
        /// Satırdaki değeri eşleşen satırın, kolonlarının toplamını, kolon başlık içeriklerine göre verir.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="KolonAdi">Satır değerinin aranacağı kolon</param>
        /// <param name="SatirDegeri">Satırdaki değer</param>
        /// <param name="KolonAdiIcerigi">Hangi metni içeren kolonların toplama dahil edileceği</param>
        /// <returns></returns>
        public static double mSatirinKolonBasliklarinaGoreToplami(this DataTable dt, string KolonAdi, string SatirDegeri, string KolonAdiIcerigi)
        {
            double toplam = 0;

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                DataRow dr = dt.Rows[r];

                if (dr[KolonAdi].ToString() == SatirDegeri)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        DataColumn col = dt.Columns[c];

                        if (col.ColumnName.Contains(KolonAdiIcerigi))
                        {
                            toplam += dt.Rows[r][c].ToString().mToDouble();
                        }
                    }

                    break;
                }
            }

            return toplam;
        }

        /// <summary>
        /// Satırların toplamını en sona bir total_*** sütunu eklerek verir.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="KolonDegeri">Kolon başlıklarının içerdiği ifade</param>
        public static void mSatirlarinToplaminiHesapla(this DataTable dt, string KolonDegeri)
        {
            string toplamKolonu = "Total_" + KolonDegeri;

            dt.Columns.Add(toplamKolonu, typeof(double));

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                DataRow dr = dt.Rows[r];

                double toplam = 0;

                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    DataColumn col = dt.Columns[c];

                    if (col.ColumnName.Contains(KolonDegeri))
                    {
                        toplam += dt.Rows[r][c].ToString().mToDouble();
                    }
                }

                dr[toplamKolonu] = toplam;
            }
        }

        public static string mToHtmlTable(this DataTable dt, bool SadeceSatirlariOlustur)
        {
            if (dt.Rows.Count == 0)
                return "";

            StringBuilder builder = new StringBuilder();

            if (!SadeceSatirlariOlustur)
            {
                builder.Append("<html>");
                builder.Append("<head>");
                builder.Append("<title>");
                builder.Append("Page-");
                builder.Append(Guid.NewGuid());
                builder.Append("</title>");
                builder.Append("</head>");
                builder.Append("<body>");
                builder.Append("<table border='1px' cellpadding='5' cellspacing='0' style='border: solid 1px Silver; font-size: x-small;'>");
            }

            builder.Append("<tr align='left' valign='top'>");
            foreach (DataColumn c in dt.Columns)
            {
                builder.Append("<td align='left' valign='top'><b>");
                builder.Append(c.ColumnName);
                builder.Append("</b></td>");
            }
            builder.Append("</tr>");
            foreach (DataRow r in dt.Rows)
            {
                builder.Append("<tr align='left' valign='top'>");
                foreach (DataColumn c in dt.Columns)
                {
                    builder.Append("<td align='left' valign='top'>");
                    builder.Append(r[c.ColumnName]);
                    builder.Append("</td>");
                }
                builder.Append("</tr>");
            }

            if (!SadeceSatirlariOlustur)
            {
                builder.Append("</table>");
                builder.Append("</body>");
                builder.Append("</html>");
            }

            return builder.ToString();
        }

        public static string mToHtmlTableV2(this DataTable dt, string tableAttribute = "")
        {
            string html = "<table " + tableAttribute + ">";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
                html += "<td>" + dt.Columns[i].ColumnName + "</td>";
            html += "</tr>";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr>";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }
    }
}
