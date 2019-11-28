using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using ExcelDataReader;
using System.Data;

namespace Rut_Task

{
    class Program
    {
        static void Main(string[] args)
        {
            // Columls List
            List<string> ListT_GK_1 = new List<string>();
            List<string> ListT_GK_2 = new List<string>();
            List<string> ListT_GK_3 = new List<string>();
            List<string> ListT_GK_4 = new List<string>();
            List<string> ListT_GK_5 = new List<string>();



            List<DataGroup> dataGroups = new List<DataGroup>();


            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source=..\..\..\excel\Rut-TaskExcel.xlsx;Extended Properties='Excel 12.0 Xml; HDR = NO;'");
            connection.Open();

            string readExcelQuery = "SELECT * FROM [Sayfa1$]";
            OleDbCommand command = new OleDbCommand(readExcelQuery, connection);
            OleDbDataReader dataReader = command.ExecuteReader();

            DataGroup dataGroup = new DataGroup("T_GK_1", "T_GK_2", "T_GK_3", "T_GK_4", "T_GK_5");

            // read from Excel
            while (dataReader.Read())
            {
                AddColumnValueToList(dataReader, dataGroup.T_GK_1, ListT_GK_1);
                AddColumnValueToList(dataReader, dataGroup.T_GK_2, ListT_GK_2);
                AddColumnValueToList(dataReader, dataGroup.T_GK_3, ListT_GK_3);
                AddColumnValueToList(dataReader, dataGroup.T_GK_4, ListT_GK_4);
                AddColumnValueToList(dataReader, dataGroup.T_GK_5, ListT_GK_5);


            }
            connection.Close();




            //Cross Datas 
            CrossJoinData(ListT_GK_1, ListT_GK_2, ListT_GK_3, ListT_GK_4, ListT_GK_5, dataGroups, dataGroup);

            //write to excel
            connection.Open();
            WriteToExcel(dataGroups, command);
            connection.Close();

            //Write to console
            WriteDataConsole(dataGroups);
            Console.ReadLine();


        }


        public static void AddColumnValueToList(OleDbDataReader dataReader, string columnName, List<string> columnList)
        {


            if (dataReader[columnName].ToString() != "")
            {
                columnList.Add(dataReader[columnName].ToString());
            }
            /*
            else
            {
                columnList.Add("");
            }
            */

        }



        private static void CrossJoinData(List<string> ListT_GK_1, List<string> ListT_GK_2, List<string> ListT_GK_3, List<string> ListT_GK_4, List<string> ListT_GK_5, List<DataGroup> dataGroups, DataGroup dataGroup)
        {
            foreach (var item1 in ListT_GK_1)
            {

                foreach (var item2 in ListT_GK_2)
                {

                    foreach (var item3 in ListT_GK_3)
                    {
                        foreach (var item4 in ListT_GK_4)
                        {
                            foreach (var item5 in ListT_GK_5)
                            {
                                dataGroup.T_GK_1 = item1;
                                dataGroup.T_GK_2 = item2;
                                dataGroup.T_GK_3 = item3;
                                dataGroup.T_GK_4 = item4;
                                dataGroup.T_GK_5 = item5;
                                dataGroups.Add(dataGroup);
                            }
                        }
                    }
                }

            }
        }
        private static void WriteToExcel(List<DataGroup> dataGroups, OleDbCommand command)
        {
            string addExcelQuery;


            foreach (DataGroup item in dataGroups)
            {
                addExcelQuery = string.Format("Insert into[Sayfa1Cross$](T_GK_1,T_GK_2,T_GK_3,T_GK_4,T_GK_5)" +
                    "values('{0}','{1}','{2}','{3}','{4}')", item.T_GK_1.ToString(), item.T_GK_2, item.T_GK_3, item.T_GK_4, item.T_GK_5);

                command.CommandText = addExcelQuery;
                command.ExecuteNonQuery();

            }
        }
        private static void WriteDataConsole(List<DataGroup> dataGroups)
        {
            Console.WriteLine("1----2----3----4----5");

            foreach (var item in dataGroups)
            {

                Console.WriteLine("{0}  {1}  {2}  {3}  {4}", item.T_GK_1, item.T_GK_2, item.T_GK_3, item.T_GK_4, item.T_GK_5);
            }
        }


    }

}
