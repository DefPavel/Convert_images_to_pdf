using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    internal static class FireService
    {
        private static readonly string StringConnection = "database=10.0.0.23:Pers;user=sysdba;password=Vtlysq~Bcgjkby2020;Charset=win1251;";

        public static async Task<Persons> getPersons(string firstName , string name , string lastName)
        {
            var persons = new Persons();
            var sql = " select first 1 s.id , s.famil , s.name , s.otch , s.date_birth , s.phone_lug , e.typ_obr ,  o.name , d.name  " +
                        " from sotr s " +
                        " left join education e on e.sotr_id = s.id " +
                        " inner join sotr_doljn sd on sd.sotr_id = s.id " +
                        " inner join doljnost d on d.id = sd.dolj_id " +
                        " inner join otdel o on o.id = d.otdel_id " +
                        $" where s.famil = '{firstName}' and s.name = '{name}' and s.otch = '{lastName.TrimEnd()}' " +
                        " order by e.is_osn desc";

            using (var connection = new FbConnection(StringConnection))
            {
                connection.Open();
                using (FbCommand cmd = new FbCommand(sql, connection))
                {
                    using (FbDataReader rd = await cmd.ExecuteReaderAsync())
                    {
                        while (await rd.ReadAsync())
                        {
                            persons = new Persons
                            {
                                FullName = $"{firstName} {name} {lastName}",
                                Birthday = rd.GetDateTime(4).ToShortDateString(),
                                Education = rd[6] != DBNull.Value ? rd[6].ToString() : null,
                                Phone = rd[5] != DBNull.Value ? rd[5].ToString() : null,
                                DepartmentPosition = $"{rd.GetString(7)}\\{rd.GetString(8)}",
                                Id = rd.GetInt32(0),
                            };
                        }
                    }
                }

            }
            return persons;
        }

        public static async Task<IEnumerable<Documents>> getDocuments(int idPerson)
        {
            var persons = new List<Documents>();
            var sql = @" select doc.name , doc.doc
                        from sotr s
                        inner join sotr_document doc on doc.sotr_id = s.id
                        where doc.typ in (2,8,7,9,22) and doc.doc is not null and s.id = " + idPerson;

            using (var connection = new FbConnection(StringConnection))
            {
                connection.Open();
                using (FbCommand cmd = new FbCommand(sql, connection))
                {
                    using (FbDataReader rd = await cmd.ExecuteReaderAsync())
                    {
                        while (await rd.ReadAsync())
                        {
                            persons.Add(new Documents
                            {
                                Name = rd.GetString(0),
                                Data = rd[1] != DBNull.Value ? (byte[])rd["doc"] : null
                            });
                        }
                    }
                }

            }
            return persons;
        }
    }
}
