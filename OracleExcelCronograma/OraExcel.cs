using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Oracle.ManagedDataAccess.Client;

namespace OracleExcelCronograma
{
    public class OraExcel
    {
        public static string strSQLcuotas = @"SELECT ROW_NUMBER() OVER (ORDER BY (NUMEROCUOTA) ASC) AS SEQNUM,
            TO_CHAR(TRUNC(FECHAVENCIMIENTO),'RRRR-MM-DD') AS FECHAVENCIMIENTO,
            TRUNC(FECHAVENCIMIENTO) - NVL((SELECT TRUNC(FECHAVENCIMIENTO) FROM PRESTAMOCUOTAS WHERE NUMEROCUOTA = (SELECT MAX(NUMEROCUOTA) AS ANTERIOR FROM PRESTAMOCUOTAS WHERE NUMEROCUOTA < PC.NUMEROCUOTA AND ESTADO <> 4 AND PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO) AND ESTADO <> 4 AND PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO), (SELECT FECHAPROGRAMACION FROM PRESTAMO WHERE PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO))
            AS DIAS,
            NVL((SELECT SALDOPRESTAMO FROM PRESTAMOCUOTAS WHERE NUMEROCUOTA = (SELECT MAX(NUMEROCUOTA) AS ANTERIOR FROM PRESTAMOCUOTAS WHERE NUMEROCUOTA < PC.NUMEROCUOTA AND ESTADO <> 4 AND PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO) AND ESTADO <> 4 AND PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO), (SELECT MONTOPRESTAMO FROM PRESTAMO WHERE PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO))
            AS SALDOINICIAL,
            NVL(AMORTIZACION, 0) + NVL(INTERES, 0) + NVL(REAJUSTE, 0) + NVL(SEGUROINTERES, 0) + NVL(MONTOSERVICIOADICIONAL, 0) + NVL(PORTES, 0)
            AS TOTALCUOTA,
            NVL(AMORTIZACION, 0) AS AMORTIZACION,
            NVL(INTERES, 0) AS INTERES,
            NVL(REAJUSTE, 0) AS DESGRAVAMEN,
            NVL(SEGUROINTERES, 0) AS SEGUROBIEN,
            NVL(MONTOSERVICIOADICIONAL, 0) AS SERVICIOADICIONAL,
            NVL(PORTES, 0) AS APORTES,
            SALDOPRESTAMO AS SALDOFINAL,
            'ESTADOCALC' AS ESTADOCALC,
            0 AS INTERESCALC,
            'REGULAR' AS TIPOCUOTA,
            TRUNC(MONTHS_BETWEEN(TRUNC(FECHAVENCIMIENTO), NVL((SELECT TRUNC(FECHAVENCIMIENTO) FROM PRESTAMOCUOTAS WHERE NUMEROCUOTA = (SELECT MAX(NUMEROCUOTA) AS ANTERIOR FROM PRESTAMOCUOTAS WHERE NUMEROCUOTA < PC.NUMEROCUOTA AND ESTADO <> 4 AND PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO) AND ESTADO <> 4 AND PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO), (SELECT FECHAPROGRAMACION FROM PRESTAMO WHERE PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO))), 1)
            AS DIFFMESES, ESTADO AS ESTADODB, INDRESTRUCTURACION,
            (SELECT SALDOPRESTAMO FROM PRESTAMO WHERE PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO)
            AS SALDOPRESTAMO,
            (SELECT TASAINTERES FROM PRESTAMODETALLE WHERE PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO AND NUMEROAMPLIACION = (SELECT MAX(NUMEROAMPLIACION) FROM PRESTAMODETALLE WHERE PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO))
            AS TEM,
            1 AS PERIODICIDAD
            FROM PRESTAMOCUOTAS PC WHERE ESTADO <> 4 AND PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO ORDER BY SEQNUM ASC";

        public static string strSQLestadoPrestamo = @"SELECT
            (CASE WHEN (SALDOPRESTAMO <= 0)
                    THEN 'CANCELADO'
                ELSE (CASE WHEN (NUMEROLINEA IS NOT NULL)
                        THEN 'LINEA'
                    ELSE (CASE WHEN (NUMEROSOLICITUDCONCESIONAL IS NOT NULL)
                            THEN 'INCREMENTO'
                        ELSE 'VIGENTE'
                        END)
                    END)
                END)
            AS ESTADO FROM PRESTAMO WHERE PERIODOSOLICITUD = :PERIODO AND NUMEROSOLICITUD = :NUMERO";

        [ExcelFunction(Description = "Consulta si Prestamo Vigente existe")]
        public static string ExistePrestamoVigenteBD(string ambiente, int periodo, int numero)
        {
            string strconn = "";
            string res = "";
            string strSQL = strSQLestadoPrestamo.Replace(":PERIODO", periodo.ToString()).Replace(":NUMERO", numero.ToString());
            try
            {
                strconn = ds.getString(ambiente);

                using (OracleConnection con = new OracleConnection(strconn))
                {
                    con.Open();
                    using (OracleCommand command = new OracleCommand(strSQL, con))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            reader.Read();
                            //Validar que el query ha devuelto alguna fila
                            if (reader.HasRows)
                            {
                                res = reader.GetString(0);
                            }
                            else
                            {
                                res = "NO EXISTE";
                            }
                        }
                        command.Dispose();
                    }
                    con.Close();
                    con.Dispose();
                }
                return res;
            }
            catch (Exception ex)
            {
                return "error:" + ex.Message + res;
            }
        }

        [ExcelFunction(Description = "Devuelve cuotas canceladas o vigentes")]
        public static string CuotasDB(string ambiente, int periodo, int numero)
        {
            DataTable dt = new DataTable();
            string strconn = "";
            string strSQL = strSQLcuotas.Replace(":PERIODO", periodo.ToString()).Replace(":NUMERO", numero.ToString());

            try
            {
                strconn = ds.getString(ambiente);
                using (OracleConnection con = new OracleConnection(strconn))
                {
                    con.Open();
                    using (OracleCommand command = new OracleCommand(strSQL, con))
                    {
                        command.ExecuteNonQuery();

                        OracleDataReader reader;

                        reader = command.ExecuteReader();
                        dt.Load(reader);
                        command.Dispose();
                    }
                    con.Close();
                    con.Dispose();
                }

                //DataTable esta en modo lectura y con numero de caracteres limitado segun query
                foreach (DataColumn col in dt.Columns)
                {
                    col.ReadOnly = false;
                    //Se analizara el tipo de cuenta mas adelante y el resultado ira en esta columna
                    if (col.ColumnName == "TIPOCUOTA")
                    {
                        col.MaxLength = 20;
                    }
                }

                if (dt.Rows.Count > 1)
                {

                    string numcuota = "0";
                    string estado = "PENDIENTE";

                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        //Prepago con cronograma pendiente
                        if (dt.Rows[i]["INDRESTRUCTURACION"].ToString() == "0")
                        {
                            numcuota = dt.Rows[i]["SEQNUM"].ToString();
                            estado = "PAGADA";
                        }
                        dt.Rows[i]["ESTADOCALC"] = estado;
                    }

                    if (numcuota == "0")
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (double.Parse(dt.Rows[i]["SALDOFINAL"].ToString()) < double.Parse(dt.Rows[i]["SALDOPRESTAMO"].ToString()))
                            {
                                estado = "PENDIENTE";
                            }
                            else
                            {
                                estado = "PAGADA";
                            }
                            dt.Rows[i]["ESTADOCALC"] = estado;
                        }
                    }
                    else
                    {
                        //Prestamo Reestructurado
                    }

                    //Periodicidad
                    List<string> difmeses = new List<string>();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["ESTADOCALC"].ToString() == "PENDIENTE")
                        {
                            difmeses.Add(dt.Rows[i]["DIFFMESES"].ToString());
                        }
                    }
                    string mindifmeses = Decimal.Round(Decimal.Parse(GetMostFrequency(difmeses, false)), 0).ToString();

                    //Analizar Tipo de Cuota
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dt.Rows[i]["INTERESCALC"] = tramoInteres(int.Parse(dt.Rows[i]["DIAS"].ToString()), Decimal.Parse(dt.Rows[i]["SALDOINICIAL"].ToString()), Decimal.Parse(dt.Rows[i]["TEM"].ToString()));
                        dt.Rows[i]["PERIODICIDAD"] = mindifmeses;

                        if (dt.Rows[i]["INDRESTRUCTURACION"].ToString() != "")
                        {
                            dt.Rows[i]["TIPOCUOTA"] = "PREPAGO";
                        }
                        else
                        {
                            if (Math.Round(Decimal.Parse(dt.Rows[i]["DIFFMESES"].ToString()) - Decimal.Parse(dt.Rows[i]["PERIODICIDAD"].ToString()), 0) >= (decimal)0.5)
                            {
                                dt.Rows[i]["TIPOCUOTA"] = "GRACIATOTAL";
                            }
                            else
                            {
                                if (Decimal.Parse(dt.Rows[i]["INTERESCALC"].ToString()) - Decimal.Parse(dt.Rows[i]["INTERES"].ToString()) <= 0 && dt.Rows[i]["AMORTIZACION"].ToString() == "0")
                                {
                                    dt.Rows[i]["TIPOCUOTA"] = "GRACIAPARCIAL";
                                }
                                else
                                {
                                    //Regular
                                }
                            }
                        }
                    }
                }

                return DataTabletoString(dt);
            }
            catch (Exception ex)
            {
                return "error:" + ex.Message;
            }
        }

        private static decimal tramoInteres(int dias, decimal saldoinicio, decimal tem)
        {
            decimal sumaInteres = 0;

            while (dias > 0)
            {
                if (dias >= 30)
                {
                    sumaInteres += 30 * Decimal.Round((saldoinicio + sumaInteres) * tem / 3000,2);
                }
                else
                {
                    sumaInteres += dias * Decimal.Round((saldoinicio + sumaInteres) * tem / 3000, 2);
                }
                dias -= 30;
            }

            return sumaInteres;
        }

        private static string GetMostFrequency(List<string> values, bool max)
        {
            var result = new Dictionary<string, int>();
            foreach (string value in values)
            {
                if (result.TryGetValue(value, out int count))
                {
                    // Increase existing value.
                    result[value] = count + 1;
                }
                else
                {
                    // New value, set to 1.
                    result.Add(value, 1);
                }
            }
            //sort list. Return most frequency.
            if (max)
            {
                var sorted = (from pair in result
                             orderby pair.Value descending, pair.Key descending
                             select pair).FirstOrDefault();
                return sorted.Key;
            }
            else
            {
                var sorted = (from pair in result
                             orderby pair.Value descending, pair.Key ascending
                             select pair).FirstOrDefault();
                return sorted.Key;
            }
        }

        private static string DataTabletoString(DataTable dt)
        {
            //列名
            string header = string.Join("|", dt.Columns.OfType<DataColumn>().Select(x => x.ColumnName));
            List<string> lstTable = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                List<string> lstRow = new List<string>();
                lstRow.Clear();
                foreach (DataColumn col in dt.Columns)
                {
                    if (!row.IsNull(col))
                    {
                        lstRow.Add(row[col].ToString());
                    }
                    else
                    {
                        lstRow.Add(string.Empty);
                    }
                }
                //行データ
                lstTable.Add(string.Join("|", lstRow));
            }
            string datas = string.Join(Environment.NewLine, lstTable);
            //列情報とデータ行を改行で区切り
            return header + Environment.NewLine + datas;
        }
    }
}
