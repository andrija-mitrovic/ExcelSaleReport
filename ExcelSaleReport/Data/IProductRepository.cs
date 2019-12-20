using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSaleReport.Data
{
    public interface IProductRepository
    {
        DataTable CompanyInformation();
        DataTable WarehouseInformation(string warehouseType);
        DataTable ProductTypeRealizationByHour(DateTime dateFrom, DateTime dateTo, string warehouseId);
        DataTable ProductTypeRealizationByDay(string warehouseId, string month);
        DataTable ProductRealizationBySupplier(string warehouseId, string month);
    }
}
