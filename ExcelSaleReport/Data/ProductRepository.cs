using Advantage.Data.Provider;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSaleReport.Data
{
    public class ProductRepository : IProductRepository
    {
        private DbConnection _connection;
        private string _query;
        public ProductRepository()
        {
            _connection = new DbConnection();
        }

        public DataTable CompanyInformation()
        {
            _query = "SELECT * FROM RO";

            return _connection.ExecuteSelectQuery(_query);
        }

        public DataTable ProductRealizationBySupplier(string warehouseId, string month)
        {
            _query = @"SELECT KOM.SIF AS ID,KOM.NAZIV AS NAME,DAY(B_NALO.DATDO) AS DTDAY,SUM(B_STAV.MPCIJ*B_STAV.KOLIC) AS RV,
                        SUM(B_STAV.MPCIJ*B_STAV.KOLIC-B_STAV.KOLIC*B_STAV.NABCJ-(B_STAV.KOLIC*(MPCIJ-AKCIZ))*M_STOPE.PROCP/(100+M_STOPE.PROCP)-B_STAV.KOLIC*R_ROBA.AKCIZ) AS DIP FROM B_NALO
                        INNER JOIN B_STAV ON B_STAV.STATU=B_NALO.STATU AND B_NALO.SIFSK=B_STAV.SIFSK AND B_NALO.BRDOK=B_STAV.BRDOK
                        INNER JOIN R_ROBA ON R_ROBA.SIFRA=B_STAV.SIFRA
                        INNER JOIN M_STOPE ON M_STOPE.STSIF=B_STAV.STSIF
                        INNER JOIN KOM ON KOM.SIF=R_ROBA.SDOBA
                        WHERE B_STAV.SIFSK=" + warehouseId +" AND MONTH(B_NALO.DATDO)=" + month + " AND B_NALO.STATU IN (8,9) AND B_NALO.PROKN=1"+
                        " GROUP BY R_ROBA.SIFRA,R_ROBA.NAZIV,KOM.SIF,DAY(B_NALO.DATDO),KOM.NAZIV"+
                        " ORDER BY DAY(B_NALO.DATDO),R_ROBA.SIFRA";

            return _connection.ExecuteSelectQuery(_query);
        }

        public DataTable ProductTypeRealizationByDay(string warehouseId, string month)
        {
            _query = @"SELECT DAY(B_NALO.DATDO) AS DTDAY, B_RADA.RADAR AS ID, B_RADA.NAZIV AS NAME, SUM(B_STAV.MPCIJ*B_STAV.KOLIC) AS REDAY FROM B_NALO
                                            INNER JOIN B_STAV ON B_NALO.SIFSK = B_STAV.SIFSK AND B_NALO.STATU = B_STAV.STATU AND B_NALO.BRDOK = B_STAV.BRDOK
                                            INNER JOIN R_ROBA ON R_ROBA.SIFRA = B_STAV.SIFRA
                                            INNER JOIN B_RADA ON B_RADA.RADAR = R_ROBA.RADAR
                                            WHERE B_NALO.STATU IN(8, 9) AND MONTH(B_NALO.DATDO) = "+month+" AND B_STAV.SIFSK = "+warehouseId+
                                            " GROUP BY DAY(B_NALO.DATDO), B_RADA.RADAR, B_RADA.NAZIV"+
                                            " ORDER BY DAY(B_NALO.DATDO)";

            return _connection.ExecuteSelectQuery(_query);
        }

        public DataTable ProductTypeRealizationByHour(DateTime dateFrom, DateTime dateTo, string warehouseId)
        {
            _query = @"SELECT LEFT(B_NALO.TIMDO,2) AS HOUR, B_RADA.RADAR AS ID, B_RADA.NAZIV AS NAME, SUM(B_STAV.MPCIJ*B_STAV.KOLIC) AS REHOUR FROM B_NALO
                                            INNER JOIN B_STAV ON B_NALO.SIFSK = B_STAV.SIFSK AND B_NALO.STATU = B_STAV.STATU AND B_NALO.BRDOK = B_STAV.BRDOK
                                            INNER JOIN R_ROBA ON R_ROBA.SIFRA = B_STAV.SIFRA
                                            INNER JOIN B_RADA ON B_RADA.RADAR = R_ROBA.RADAR
                                            WHERE B_NALO.STATU IN(8, 9) AND B_NALO.DATDO >='" + Convert.ToDateTime(dateFrom).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture) 
                                            + "' AND B_NALO.DATDO<='" + Convert.ToDateTime(dateTo).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture) + 
                                            "'  AND B_STAV.SIFSK ="+warehouseId+
                                            " GROUP BY LEFT(B_NALO.TIMDO, 2), B_RADA.RADAR, B_RADA.NAZIV" +
                                            " ORDER BY LEFT(B_NALO.TIMDO, 2)";

            return _connection.ExecuteSelectQuery(_query);
        }

        public DataTable WarehouseInformation(string warehouseType)
        {
            _query = "SELECT SIF AS ID,NAZIV AS NAME FROM KOM WHERE sif > 50000 AND POZNABR LIKE '" + warehouseType + "%' ORDER BY SIF";

            return _connection.ExecuteSelectQuery(_query);
        }
    }
}
