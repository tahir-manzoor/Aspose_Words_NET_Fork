using System;

namespace ApiExamples.TestData
{
}

namespace ApiExamples.TestData
{
    internal static class TestTables
    {
        internal static DataSet AddClientsTestData()
        {
            Random rnd = new Random();
            DataSet ds = new DataSet();

            int j = 1;

            for (int i = 1; i <= 10; i++)
            {
                ds.Clients.Rows.Add(i, "Name " + i);
            }

            for (int i = 1; i <= 3; i++)
            {
                for (int y = j; y <= j + 2; y++)
                {
                    ds.Contracts.Rows.Add(y, y, i, rnd.Next(2000, 100000));
                }

                j = j + 3;
            }

            for (int i = 1; i <= 3; i++)
            {
                ds.Managers.Rows.Add(i, "Name " + i, rnd.Next(20, 50));
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
