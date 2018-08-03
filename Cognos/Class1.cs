using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Cognos
{
    public class Class1
    {

        SqlConnection cnn;
        public List<Tuple<Entity, Tuple<int, int>>> tasks = new List<Tuple<Entity, Tuple<int, int>>>();
        public Tuple<int, int> _cellAddress = null;
        public Tuple<int, int> cellAddress
        {
            get
            {
                if (_cellAddress != null)
                {
                    return _cellAddress;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (_cellAddress == null || _cellAddress.Item1 != value.Item1 || _cellAddress.Item2 != value.Item2)
                {
                    _cellAddress = value;
                    Debug.Print("Prime and execute");
                }
            }
        }

        public void addTask(Tuple<int, int> workspace, Entity entity)
        {
            if (this.tasks != null && this.tasks.Count > 0)
            {
                int oldData = -1;
                try
                {
                    oldData = this.tasks.FindIndex(x =>
                    x.Item2.Item1 == workspace.Item1 && x.Item2.Item2 == workspace.Item2);
                }
                catch (Exception ex)
                {
                    Debug.Print(ex.Message.ToString());
                }
                if (oldData != -1)
                {
                    this.tasks.RemoveAt(Convert.ToInt32(oldData));
                }
                this.tasks.Add(new Tuple<Entity, Tuple<int, int>>(entity, workspace));
            }
            else
            {
                this.tasks.Add(new Tuple<Entity, Tuple<int, int>>(entity, workspace));
            }
        }

        

        public void connectToDB()
        {
            string connetionString = null;

            
            cnn = new SqlConnection(connetionString);

            try
            {
                cnn.Open();
                Debug.Print("Connection Open!");
                // cnn.Close();
            }
            catch (Exception ex)
            {
                Debug.Print($"Can not open connection! Error: {ex.Message.ToString()}");
            }
            //cnn.Close();
        }

        public string getBeloppFromDB(Entity entity)
        {
            
            command.Parameters.AddWithValue("@name", "*");
            string stringToReturn = null;
            // int result = command.ExecuteNonQuery();
            using (SqlDataReader reader = command.ExecuteReader())
            {

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        var e = new Entity(reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetString(3), reader.GetString(4), reader.GetString(5), reader.GetString(6), reader.GetString(7), reader.GetString(8), reader.GetString(9), reader.GetString(10), reader.GetInt32(11), reader.GetString(12), reader.GetString(13), reader.GetString(14), reader.GetString(15), reader.GetDecimal(16), reader.GetDecimal(17), reader.GetString(18), reader.GetInt32(19), reader.GetInt32(20), reader.GetInt16(21));
                        //Debug.Print(e.belopp.ToString());
                        stringToReturn = e.belopp.ToString();
                    }
                }
                else
                {
                    Console.WriteLine("No rows found.");
                    stringToReturn = "No entity found";
                }
                reader.Close();
            }
            return stringToReturn;
        }

        public void closeConnection()
        {
            cnn.Close();
        }
    }
}
