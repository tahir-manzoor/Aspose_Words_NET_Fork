using System;
using System.Data;

namespace ApiExamples.TestData
{
    internal static class TestTables
    {
        internal static DataSet AddClientsTestData()
        {
            var rnd = new Random();
            DataSet ds = new DataSet();

            for (var i = 1; i <= 10; i++)
            {
                ds.Clients.Rows.Add(i, "Name " + i);
            }

            for (var i = 1; i <= 3; i++)
            {
                ds.Contracts.Rows.Add(i, i, 1, rnd.Next(2000, 100000));
            }

            for (var i = 4; i <= 6; i++)
            {
                ds.Contracts.Rows.Add(i, i, 2, rnd.Next(2000, 100000));
            }

            for (var i = 7; i <= 9; i++)
            {
                ds.Contracts.Rows.Add(i, i, 3, rnd.Next(2000, 100000));
            }

            for (var i = 1; i <= 3; i++)
            {
                ds.Managers.Rows.Add(i, "Name " + i);
            }

            return ds;
        }
    }
}

namespace ApiExamples.TestData
{


    partial class DataSet
    {
    }
}
