using Microsoft.Extensions.Configuration;
using PersonInfoWebAPIWPF.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace PersonInfoWebAPIWPF.Data
{
    public class PersonAccess
    {
        private readonly string _msSqlConnectionString;
        private readonly string _msAccessConnectionString;
        private readonly string _azureSqlConnectionString;

        private SqlConnection SqlConnectionBuilder(string conString)
        {
            return new SqlConnection(conString);
        }

        private OleDbConnection OleDbConnectionBuilder(string conString)
        {
            return new OleDbConnection();
        }

        private IConfigurationRoot GetConfiguration()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            return builder.Build();
        }

        public PersonAccess()
        {
            var configuration = GetConfiguration();
            var conStrings = configuration.GetSection("ConnectionStrings");
            Console.WriteLine(Environment.CurrentDirectory);
            string access_db_file_name = "naba_db.accdb";
            string access_db_path = System.IO.Path.Combine(Environment.CurrentDirectory, @"..\PersonInfoWPFApp\Files\", access_db_file_name);
            //_msAccessConnectionString = conStrings.GetSection("msAccessConnection").Value;
            _msAccessConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Provider=Microsoft.ACE.OLEDB.12.0;Data Source={access_db_path};Persist Security Info=True;Persist Security Info=False;";
            _msSqlConnectionString = conStrings.GetSection("msSqlConnection").Value;
            _azureSqlConnectionString = conStrings.GetSection("azureSqlConnection").Value;
        }

        private Person MapToPerson(SqlDataReader reader)
        {
            return new Person()
            {
                Id = (int)reader["Id"],
                FirstName = reader["FirstName"].ToString(),
                LastName = reader["LastName"].ToString()
            };
        }

        public async Task<List<Person>> GetAll(string dbSelection)
        {
            if (dbSelection == "msAccessConnection")
            {
                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = _msAccessConnectionString;
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = $"SELECT * FROM Person;";
                OleDbDataReader reader = command.ExecuteReader();
                List<Person> people = new List<Person>();
                try
                {
                    while (reader.Read())
                    {
                        Person person = new Person
                        {
                            Id = (int)reader["Id"],
                            FirstName = (string)reader["FirstName"],
                            LastName = (string)reader["LastName"] // or whatever the type should be
                        };
                        people.Add(person);
                    }
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                    return new List<Person> { new Person() };
                }
                connection.Close();
                return people;
            }
            if (dbSelection == "msSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_msSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("GetAllPerson", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        var response = new List<Person>();
                        await cnn.OpenAsync();

                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                response.Add(MapToPerson(reader));
                            }
                        }
                        return response;
                    }
                }
            }
            if (dbSelection == "azureSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_azureSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("GetAllPerson", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        var response = new List<Person>();
                        await cnn.OpenAsync();

                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                response.Add(MapToPerson(reader));
                            }
                        }
                        return response;
                    }
                }
            }
            return null;
        }

        public async Task<Person> GetById(int Id, string dbSelection)
        {
            if (dbSelection == "msAccessConnection")
            {
                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = _msAccessConnectionString;
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = $"SELECT * FROM Person WHERE Id={Id};";
                OleDbDataReader reader = command.ExecuteReader();
                Person person = new Person();
                try
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            person = new Person
                            {
                                Id = (int)reader["Id"],
                                FirstName = (string)reader["FirstName"],
                                LastName = (string)reader["LastName"]
                            };
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return new Person();
                }
                connection.Close();
                return person;
            }
            if (dbSelection == "msSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_msSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("GetPersonById", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@Id", Id));
                        Person person = null;
                        await cnn.OpenAsync();

                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                person = MapToPerson(reader);
                            }
                        }
                        return person;
                    }
                }
            }
            if (dbSelection == "azureSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_azureSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("GetPersonById", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@Id", Id));
                        Person response = null;
                        await cnn.OpenAsync();

                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                response = MapToPerson(reader);
                            }
                        }
                        return response;
                    }
                }
            }
            return null;
        }

        public async Task Insert(Person person, string dbSelection)
        {
            if (dbSelection == "msAccessConnection")
            {
                try
                {
                    OleDbConnection connection = new OleDbConnection();
                    connection.ConnectionString = _msAccessConnectionString;
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;
                    command.CommandText = $"INSERT INTO Person (FirstName, LastName) VALUES ('{person.FirstName}', '{person.LastName}');";
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return;
                }
            }
            if (dbSelection == "msSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_msSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("InsertPerson", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@FirstName", person.FirstName));
                        cmd.Parameters.Add(new SqlParameter("@LastName", person.LastName));
                        await cnn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                        return;
                    }
                }
            }
            if (dbSelection == "azureSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_azureSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("InsertPerson", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@FirstName", person.FirstName));
                        cmd.Parameters.Add(new SqlParameter("@LastName", person.LastName));
                        await cnn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                        return;
                    }
                }
            }
            return;
        }

        public async Task UpdateById(int Id, Person person, string dbSelection)
        {
            if (dbSelection == "msAccessConnection")
            {
                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = _msAccessConnectionString;
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = $"UPDATE Person SET FirstName='{person.FirstName}', LastName='{person.LastName}' WHERE Id={person.Id};";
                command.ExecuteNonQuery();
                connection.Close();
            }
            if (dbSelection == "msSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_msSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("UpdatePersonById", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@Id", Id));
                        cmd.Parameters.Add(new SqlParameter("@FirstName", person.FirstName));
                        cmd.Parameters.Add(new SqlParameter("@LastName", person.LastName));
                        await cnn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                        return;
                    }
                }
            }
            if (dbSelection == "azureSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_azureSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("UpdatePersonById", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@Id", Id));
                        cmd.Parameters.Add(new SqlParameter("@FirstName", person.FirstName));
                        cmd.Parameters.Add(new SqlParameter("@LastName", person.LastName));
                        await cnn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                        return;
                    }
                }
            }
            return;
        }

        public async Task DeleteById(int Id, string dbSelection)
        {
            if (dbSelection == "msAccessConnection")
            {
                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = _msAccessConnectionString;
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = $"DELETE FROM Person WHERE Id={Id};";
                command.ExecuteNonQuery();
                connection.Close();
            }
            if (dbSelection == "msSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_msSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("DeletePersonById", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@Id", Id));
                        await cnn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                        return;
                    }
                }
            }
            if (dbSelection == "azureSqlConnection")
            {
                using (SqlConnection cnn = new SqlConnection(_azureSqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("DeletePersonById", cnn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@Id", Id));
                        await cnn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                        return;
                    }
                }
            }
            return;
        }
    }
}