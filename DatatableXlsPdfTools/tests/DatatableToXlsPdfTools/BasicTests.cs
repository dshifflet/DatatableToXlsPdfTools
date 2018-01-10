using System.Configuration;
using System.Data;
using System.IO;
using System.Threading;
using DatatableXlsPdfTools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Oracle.ManagedDataAccess.Client;

namespace DatatableToXlsPdfTools.Tests
{
    [TestClass]
    public class BasicTests
    {
        [TestMethod]
        public void CanConvertToPdfAndXls()
        {
            using (var cn = 
                new OracleConnection(ConfigurationManager.ConnectionStrings["NORTHWIND"].ConnectionString))
            {
                cn.Open();
                using (var cmd = cn.CreateCommand())
                {
                    cmd.CommandText = "select * from northwind.categories";
                    using (var reader = cmd.ExecuteReader())
                    {
                        var dataTable = new DataTable();
                        dataTable.Load(reader);

                        var files = new[]
                        {
                            new FileInfo("test.pdf"),
                            new FileInfo("test.xlsx")
                        };

                        foreach (var file in files)
                        {
                            if (file.Exists)
                            {
                                file.Delete();
                            }
                            
                            DataTableToXlsPdf.ToFile(dataTable, file);
                            Thread.Sleep(5000);
                            Assert.IsTrue(file.Exists);
                        }
                    }
                }
            }
        }
    }
}
