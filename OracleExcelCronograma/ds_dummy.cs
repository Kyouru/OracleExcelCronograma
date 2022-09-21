using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OracleExcelCronograma
{
    public static class ds
    {
        public static string strConnectionDesa = "Data source=DESARROLLO_NEW;CredencialDesa;";
        public static string strConnectionQA = "Data source=DESARROLLO_QA;CredencialQA;";
        public static string strConnectionProd = "Data source=BDPACIFICO;CredencialProd;";

        public static string getString(string ambiente)
        {
            string strconn = "";
            switch (ambiente)
            {
                case "QA":
                    strconn = strConnectionQA.Replace("DESARROLLO_QA", QA);
                    break;
                case "DESA":
                    strconn = strConnectionDesa.Replace("DESARROLLO_NEW", DESA);
                    break;
                case "PROD":
                    strconn = strConnectionProd.Replace("BDPACIFICO", PROD);
                    break;
                default:
                    strconn = strConnectionDesa.Replace("DESARROLLO_NEW", DESA);
                    break;
            }
            return strconn;
        }


        public static string DESA = "DataSourceDesa";
        public static string QA = "DataSourceQA";
        public static string PROD = "DataSourceProd";

    }
}
