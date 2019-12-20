using ExcelSaleReport.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelSaleReport
{
    public class Reports
    {
        private IProductRepository _repo;
        private List<string> _warehouseType;
        private readonly string _title;
        private int sheetNumber;
        private DataTable _dtExcel;
        private DataTable _dtWarehouse;
        private DataTable _dtCompany;
        private ExcelDocument _excel;
        private const int hours = 24;
        private const int days = 31;

        public Reports(IProductRepository repo)
        {
            _repo = repo;
            _excel = new ExcelDocument();
            _warehouseType = new List<string>();
            _warehouseType.Add("SN");
            _warehouseType.Add("CN");

            _dtCompany = _repo.CompanyInformation();
            if(_dtCompany.Rows.Count != 0)
                _title = _dtCompany.Rows[0]["ronaz"].ToString().ToUpper().Trim(' ') + " " + _dtCompany.Rows[0]["p_god"].ToString();
        }

        public void GetProductTypeRealizationByHour(DateTime dateFrom, DateTime dateTo)
        {
            try
            {
                if (_dtCompany.Rows.Count == 0)
                {
                    MessageBox.Show("There is no company information in database...", "Attention");
                    return;
                }

                _dtExcel = new DataTable();
                _dtExcel.Columns.Add("Id", typeof(Int32));
                _dtExcel.Columns.Add("Name", typeof(String));
                for (int j = 0; j < hours; j++)
                    _dtExcel.Columns.Add(j.ToString(), typeof(Decimal));
                _dtExcel.Columns.Add("Total", typeof(Decimal));
                _dtExcel.PrimaryKey = new DataColumn[] { _dtExcel.Columns["Id"] };

                string header = "REALIZATION BY HOUR FOR PERIOD " + dateFrom.ToString("dd.MM.yyyy")
                    + "-" + dateTo.ToString("dd.MM.yyyy");
                sheetNumber = 1;
                for (int i = 0; i < _warehouseType.Count; i++)
                {
                    _dtWarehouse = _repo.WarehouseInformation(_warehouseType[i]);

                    foreach (DataRow rowWarehouse in _dtWarehouse.Rows)
                    {
                        DataTable dtProducts = _repo.ProductTypeRealizationByHour(dateFrom,dateTo,rowWarehouse["Id"].ToString());

                        if (dtProducts.Rows.Count != 0)
                        {
                            foreach (DataRow row in dtProducts.Rows)
                            {
                                var frow = _dtExcel.Rows.Find(new Object[] { Int32.Parse(row["Id"].ToString()) });
                                if (frow == null)
                                {
                                    _dtExcel.Rows.Add(new Object[] { Int32.Parse(row["Id"].ToString()), row["Name"].ToString().Trim(' '),
                                0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 });
                                    frow = _dtExcel.Rows.Find(new Object[] { Int32.Parse(row["Id"].ToString()) });
                                    int hour = Convert.ToInt32(row["hour"]);
                                    frow[hour.ToString()] = Convert.ToDecimal(frow[hour.ToString()]) + Convert.ToDecimal(row["rehour"]);
                                }
                                else
                                {
                                    int hour = Convert.ToInt32(row["hour"]);
                                    frow[hour.ToString()] = Convert.ToDecimal(frow[hour.ToString()]) + Convert.ToDecimal(row["rehour"]);
                                }
                            }

                            foreach (DataRow row in _dtExcel.Rows)
                            {
                                double sum = 0;
                                for (int j = 0; j < hours; j++)
                                {
                                    sum += Convert.ToDouble(row[j.ToString()]);
                                }
                                row["total"] = sum;
                            }


                            _excel.CreateSheet(rowWarehouse["name"].ToString().Trim(' '), sheetNumber, _title, header, _dtExcel);
                            sheetNumber++;
                            _dtExcel.Clear();
                            }
                        }
                 }

                if (sheetNumber != 1)
                    _excel.SaveExcelDoc("Realization_By_Hour");
                else
                    MessageBox.Show("There is no data for above period!");
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error - RealizationByHour! \n Exception: "+ ex.Message);
            }
        }


        public void GetProductTypeRealizationByDay(int month)
        {
            try
            {
                if (_dtCompany.Rows.Count == 0)
                {
                    MessageBox.Show("There is no company information in database...", "Attention");
                    return;
                }

                _dtExcel = new DataTable();
                _dtExcel.Columns.Add("Id", typeof(Int32));
                _dtExcel.Columns.Add("Name", typeof(String));
                for (int j = 1; j <= days; j++)
                    _dtExcel.Columns.Add(j.ToString(), typeof(decimal));
                _dtExcel.Columns.Add("Total", typeof(decimal));
                _dtExcel.PrimaryKey = new DataColumn[] { _dtExcel.Columns["Id"] };

                string header = "REALIZATION BY DAY FOR " + month + " MONTH";
                sheetNumber = 1;
                for (int i = 0; i < _warehouseType.Count; i++)
                {
                    _dtWarehouse = _repo.WarehouseInformation(_warehouseType[i]);
                    foreach (DataRow rowWarehouse in _dtWarehouse.Rows)
                    {
                        DataTable dtProducts = _repo.ProductTypeRealizationByDay(rowWarehouse["ID"].ToString(), month.ToString());

                        if (dtProducts.Rows.Count != 0)
                        {
                            foreach (DataRow row in dtProducts.Rows)
                            {
                                var frow = _dtExcel.Rows.Find(new Object[] { Int32.Parse(row["Id"].ToString()) });
                                if (frow == null)
                                {
                                    _dtExcel.Rows.Add(new Object[] { Int32.Parse(row["Id"].ToString()), row["Name"].ToString().Trim(' '), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 });
                                    frow = _dtExcel.Rows.Find(new Object[] { Int32.Parse(row["Id"].ToString()) });
                                    int dan = Convert.ToInt32(row["dtday"]);
                                    frow[dan.ToString()] = Convert.ToDecimal(frow[dan.ToString()]) + Convert.ToDecimal(row["reday"]);
                                }
                                else
                                {
                                    int dan = Convert.ToInt32(row["dtday"]);
                                    frow[dan.ToString()] = Convert.ToDecimal(frow[dan.ToString()]) + Convert.ToDecimal(row["reday"]);
                                }
                            }

                            foreach (DataRow row in _dtExcel.Rows)
                            {
                                double sum = 0;
                                for (int j = 1; j <= days; j++)
                                {
                                    sum += Convert.ToDouble(row[j.ToString()]);
                                }
                                row["total"] = sum;
                            }

                            _excel.CreateSheet(rowWarehouse["Name"].ToString().Trim(' '), sheetNumber, _title, header, _dtExcel);
                            sheetNumber++;
                            _dtExcel.Clear();
                        }
                    }
                }

                if (sheetNumber != 1)
                    _excel.SaveExcelDoc("Realization_By_Day");
                else
                    MessageBox.Show("There is no data for above period!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error - RealizationByMonth! \n Exception: " + ex.Message);
            }
        }

        public void GetProductBySupplier(int month)
        {
            if (_dtCompany.Rows.Count == 0)
            {
                MessageBox.Show("There is no company information in database...", "Attention");
                return;
            }
            _dtExcel = new DataTable();
            _dtExcel.Columns.Add("ID", typeof(Int32));
            _dtExcel.Columns.Add("Name", typeof(String));

            for (int j = 1; j <= days; j++)
            {
                _dtExcel.Columns.Add("% DIP" + j.ToString(), typeof(decimal));
                _dtExcel.Columns.Add("DIP" + j.ToString(), typeof(decimal));
                _dtExcel.Columns.Add("RV" + j.ToString(), typeof(decimal));
            }

            _dtExcel.Columns.Add("%DIPTotal", typeof(decimal));
            _dtExcel.Columns.Add("DIPTotal", typeof(decimal));
            _dtExcel.Columns.Add("RVTotal", typeof(decimal));
            _dtExcel.PrimaryKey = new DataColumn[] { _dtExcel.Columns["ID"] };

            string header = "SUPPLIER REALIZATION BY DAY FOR " + month + " MONTH";
            sheetNumber = 1;
            for (int i = 0; i < _warehouseType.Count; i++)
            {
                _dtWarehouse = _repo.WarehouseInformation(_warehouseType[i]);
                foreach (DataRow rowWarehouse in _dtWarehouse.Rows)
                {
                    DataTable dtProducts = _repo.ProductRealizationBySupplier(rowWarehouse["ID"].ToString(), month.ToString());

                    if (dtProducts.Rows.Count != 0)
                    {
                        foreach (DataRow row in dtProducts.Rows)
                        {
                            var frow = _dtExcel.Rows.Find(new Object[] { Int32.Parse(row["ID"].ToString()) });
                            if (frow == null)
                            {
                                _dtExcel.Rows.Add(new Object[] { Int32.Parse(row["ID"].ToString()), row["Name"].ToString().Trim(' '),
                                 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                                 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                                 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0});

                                frow = _dtExcel.Rows.Find(new Object[] { Int32.Parse(row["ID"].ToString()) });
                                int dan = Convert.ToInt32(row["DTDAY"]);
                                frow[("RV" + dan).ToString()] = Convert.ToDecimal(frow[("RV" + dan).ToString()]) + Convert.ToDecimal(row["RV"]);
                                frow[("DIP" + dan).ToString()] = Convert.ToDecimal(frow[("DIP" + dan).ToString()]) + Convert.ToDecimal(row["DIP"]);
                            }
                            else
                            {
                                int dan = Convert.ToInt32(row["DTDAY"]);
                                frow[("RV" + dan).ToString()] = Convert.ToDecimal(frow[("RV" + dan).ToString()]) + Convert.ToDecimal(row["RV"]);
                                frow[("DIP" + dan).ToString()] = Convert.ToDecimal(frow[("DIP" + dan).ToString()]) + Convert.ToDecimal(row["DIP"]);
                            }
                        }

                        for (int j = 1; j < _dtExcel.Columns.Count / 3; j++)
                        {
                            int sum = Convert.ToInt32(_dtExcel.Compute("SUM(RV" + j + ")", string.Empty));
                            foreach (DataRow row in _dtExcel.Rows)
                            {
                                if (sum != 0)
                                    row["% DIP" + j] = Convert.ToDecimal(row[("DIP" + j).ToString()]) / sum * 100;
                            }
                        }

                        foreach (DataRow row in _dtExcel.Rows)
                        {
                            double sumDIP = 0;
                            double sumRV = 0;
                            for (int j = 1; j <= days; j++)
                            {
                                sumDIP += Convert.ToDouble(row["DIP" + j.ToString()]);
                                sumRV += Convert.ToDouble(row["RV" + j.ToString()]);
                            }
                            row["DIPTotal"] = sumDIP;
                            row["RVTotal"] = sumRV;
                        }

                        
                        int sumDIPCol = Convert.ToInt32(_dtExcel.Compute("SUM(DIPTotal)", string.Empty));
                        foreach (DataRow row in _dtExcel.Rows)
                        {
                            if (sumDIPCol != 0)
                                row["%DIPTotal"] = Convert.ToDecimal(row["DIPTotal"]) / sumDIPCol * 100;
                        }

                        _excel.CreateSheetWithMultiLayerColumn(rowWarehouse["NAME"].ToString().Trim(' '), sheetNumber, _title, header, _dtExcel);
                        sheetNumber++;
                        _dtExcel.Clear();
                    }
                }
            }

            if (sheetNumber != 1)
                _excel.SaveExcelDoc("Realization_By_Supplier");
            else
                MessageBox.Show("There is no data for above period!");
        }
    }
}
