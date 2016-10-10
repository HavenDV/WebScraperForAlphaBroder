using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BigCommerce
{
    static class BigCommerceUtilities
    {
        //Color string
        //"  Rule,,\"[CS]Color=Black:#000000\",,,,,,,,[ADD]{add_price},,,,,,,,,,,Y,Y,,,,,,,{image},,N,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,N,,,,,,,,,,,,,,,"
        //Size string
        //"  Rule",,\"[CS]Color=Black:#000000,[RT]Size=XXL\",,,,,,,,[ADD]4,,,,,,,[ADD]0.01,,,,Y,Y,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,N,,,,,,,,,,,,,,,
        //No size
        //"  Rule",,\"[CS]Color=Black:#000000,[RT]Size=XXXL\",,,,,,,,,,,,,,,[ADD]0.02,,,,N,Y,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,N,,,,,,,,,,,,,,,

        public static string ProductFormat(string name, string desc, string prefix, string id, string brand, string price, string width, IList<string> images, string category, string subcategory, string collection = "")
        {

            var template = "Product,,\"{0}\",P,{12}{11},,{13},\"{13} {0}\",Right,\"<p><span>{1}</span></p>\",{2},0.00,0.00,0.00,0.00,N,,{14},0.0000,0.0000,0.0000,Y,Y,,none,0,0,\"Shop/{8}/{9}/{10}\",,{3},,Y,0,,{4},,,,,{5},,,,,{6},,,,,{7},,,,,,,,,,,New,N,N,\"Delivery Date\",N,,,0,\"Non - Taxable Products\",,N,,,,,,,,,,,,,N,,";
            return string.Format(template,
                name,
                desc.Trim(' ', '\n', '\r').Replace("\r", "").Replace("\n", ""),
                price,
                images.Count > 0 ? images[0] : "",
                images.Count > 1 ? images[1] : "",
                images.Count > 2 ? images[2] : "",
                images.Count > 3 ? images[3] : "",
                images.Count > 4 ? images[4] : "",
                category,
                subcategory,
                collection, id, prefix, brand, width);
        }

        public static string ColorFormat(string color, IList<string> images, string preview)
        {
            return string.Format(
                "  Rule,,\"[CS]Color={0}:{6}\",,,,,,,,,,,,,,,,,,,Y,Y,,,,,,,{1},,N,,,{2},,,,,{3},,,,,{4},,,,,{5},,,,,,,,,,,,,,,,,,,,,N,,,,,,,,,,,,,,,",
                color,
                images.Count > 0 ? images[0] : "",
                images.Count > 1 ? images[1] : "",
                images.Count > 2 ? images[2] : "",
                images.Count > 3 ? images[3] : "",
                images.Count > 4 ? images[4] : "",
                preview);
        }

        public static string GetPrefixForValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : "[FIXED]";
        }

        public static string SizeFormat(string size, string price, string weight)
        {
            return string.Format(
                "  Rule,,\"[RT]Size={0}\",,,,,,,,{3}{1},,,,,,,{4}{2},,,,Y,Y,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,N,,,,,,,,,,,,,,,", 
                size, price, weight,
                GetPrefixForValue(price),
                GetPrefixForValue(weight));
        }

        public static string ColorSizeFormat(string color, string size, string price, string weight, string previewImage)
        {
            return string.Format(
                "  Rule,,\"[CS]Color={0}:{4},[RT]Size={1}\",,,,,,,,{5}{2},,,,,,,{6}{3},,,,Y,Y,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,N,,,,,,,,,,,,,,,",
                color, size, price, weight, previewImage,
                GetPrefixForValue(price),
                GetPrefixForValue(weight)
                );
        }

        public static string Header {
            get {
                return "\"Item Type\",\"Product ID\",\"Product Name\",\"Product Type\",\"Product Code/SKU\",\"Bin Picking Number\",\"Brand Name\",\"Option Set\",\"Option Set Align\",\"Product Description\",\"Price\",\"Cost Price\",\"Retail Price\",\"Sale Price\",\"Fixed Shipping Cost\",\"Free Shipping\",\"Product Warranty\",\"Product Weight\",\"Product Width\",\"Product Height\",\"Product Depth\",\"Allow Purchases?\",\"Product Visible?\",\"Product Availability\",\"Track Inventory\",\"Current Stock Level\",\"Low Stock Level\",\"Category\",\"Product Image ID - 1\",\"Product Image File - 1\",\"Product Image Description - 1\",\"Product Image Is Thumbnail - 1\",\"Product Image Sort - 1\",\"Product Image ID - 2\",\"Product Image File - 2\",\"Product Image Description - 2\",\"Product Image Is Thumbnail - 2\",\"Product Image Sort - 2\",\"Product Image ID - 3\",\"Product Image File - 3\",\"Product Image Description - 3\",\"Product Image Is Thumbnail - 3\",\"Product Image Sort - 3\",\"Product Image ID - 4\",\"Product Image File - 4\",\"Product Image Description - 4\",\"Product Image Is Thumbnail - 4\",\"Product Image Sort - 4\",\"Product Image ID - 5\",\"Product Image File - 5\",\"Product Image Description - 5\",\"Product Image Is Thumbnail - 5\",\"Product Image Sort - 5\",\"Search Keywords\",\"Page Title\",\"Meta Keywords\",\"Meta Description\",\"MYOB Asset Acct\",\"MYOB Income Acct\",\"MYOB Expense Acct\",\"Product Condition\",\"Show Product Condition?\",\"Event Date Required?\",\"Event Date Name\",\"Event Date Is Limited?\",\"Event Date Start Date\",\"Event Date End Date\",\"Sort Order\",\"Product Tax Class\",\"Product UPC/EAN\",\"Stop Processing Rules\",\"Product URL\",\"Redirect Old URL?\",\"GPS Global Trade Item Number\",\"GPS Manufacturer Part Number\",\"GPS Gender\",\"GPS Age Group\",\"GPS Color\",\"GPS Size\",\"GPS Material\",\"GPS Pattern\",\"GPS Item Group ID\",\"GPS Category\",\"GPS Enabled\",\"Avalara Product Tax Code\",\"Product Custom Fields\"";
            }
        }

        public static StreamWriter CreateImportCSVFile(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var writer = File.CreateText(path);
            writer.WriteLine(Header);
            return writer;
        }
    }
}
