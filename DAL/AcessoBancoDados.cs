using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;

namespace DAL
{
    public class AcessoBancoDados
    {
        //Conexão com o Banco de Dados
        private SqlConnection CriarConexao()
        {
            //return new SqlConnection("Data Source=.\\SQLEXPRESS;Initial Catalog=WebScrapingImpressoras;Integrated Security=True");
            return new SqlConnection("Data Source =srvsao008\\srvsa008; Initial Catalog = WebScrapingImpressoras; Integrated Security = True");
        }

        //Parametros que vão para o banco
        private SqlParameterCollection sqlParameterCollection = new SqlCommand().Parameters;

        public void LimparParametros()
        {
            sqlParameterCollection.Clear();
        }

        public void AdicionarParametros(string nomeParametro, object valorParametro)
        {
            sqlParameterCollection.Add(new SqlParameter(nomeParametro, valorParametro));
        }

        //Persistência - Inserir,  Alterar, Excluir
        public object ExecutarManipulacao(CommandType commandType, string nomeProcedureOuTextoSQL)
        {
            try
            {
                //Criar conexão
                SqlConnection sqlConnection = CriarConexao();
                sqlConnection.Open();

                //Criar o comando que levara informação para o banco.
                SqlCommand sqlCommand = sqlConnection.CreateCommand();

                //Colocando as informações dentro do comando.
                sqlCommand.CommandType = commandType;
                sqlCommand.CommandText = nomeProcedureOuTextoSQL;
                sqlCommand.CommandTimeout = 7200; //Tempo de execução máximo - Em Segundos --- 2Horas.

                //Adicionar os parâmetros no comando
                foreach (SqlParameter sqlParameter in sqlParameterCollection)
                {
                    sqlCommand.Parameters.Add(new SqlParameter(sqlParameter.ParameterName, sqlParameter.Value));
                }

                //Executar o comando, mandar o comando ir até o banco de dados.
                return sqlCommand.ExecuteScalar();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        //Consultar registros do banco de dados
        public DataTable ExecutarConsulta(CommandType commandType, string nomeProcedureOuTextoSQL)
        {
            try
            {
                SqlConnection sqlConnection = CriarConexao();
                sqlConnection.Open();
                SqlCommand sqlCommand = sqlConnection.CreateCommand();

                sqlCommand.CommandType = commandType;
                sqlCommand.CommandText = nomeProcedureOuTextoSQL;
                sqlCommand.CommandTimeout = 7200; //Em Segundos --- 2Horas.

                foreach (SqlParameter sqlParameter in sqlParameterCollection)
                {
                    sqlCommand.Parameters.Add(new SqlParameter(sqlParameter.ParameterName, sqlParameter.Value));
                }

                SqlDataAdapter da = new SqlDataAdapter(sqlCommand);
                DataTable dt = new DataTable();
                da.Fill(dt);


                return dt;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public int VerificarLogin(CommandType commandType, string nomeProcedureOuTextoSQL)
        {
            var login = 0;

            try
            {
                SqlConnection sqlConnection = CriarConexao();
                sqlConnection.Open();
                SqlCommand sqlCommand = sqlConnection.CreateCommand();

                sqlCommand.CommandType = commandType;
                sqlCommand.CommandText = nomeProcedureOuTextoSQL;
                sqlCommand.CommandTimeout = 7200; //Em Segundos --- 2Horas.

                foreach (SqlParameter sqlParameter in sqlParameterCollection)
                {
                    sqlCommand.Parameters.Add(new SqlParameter(sqlParameter.ParameterName, sqlParameter.Value)).Direction = ParameterDirection.Input;
                }
                sqlCommand.Parameters.Add("@Retorno", SqlDbType.Int).Direction = ParameterDirection.Output;
                SqlParameter retorno = sqlCommand.Parameters.Add("@Retorno", SqlDbType.Int);
                retorno.Direction = ParameterDirection.ReturnValue;

                SqlDataReader dr = sqlCommand.ExecuteReader();

                login = Convert.ToInt32(sqlCommand.Parameters["@Retorno"].Value.ToString());

                return login;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public int VerificarSolicitacaoExiste(CommandType commandType, string nomeProcedureOuTextoSQL)
        {
            var existe = 0;

            try
            {
                SqlConnection sqlConnection = CriarConexao();
                sqlConnection.Open();
                SqlCommand sqlCommand = sqlConnection.CreateCommand();

                sqlCommand.CommandType = commandType;
                sqlCommand.CommandText = nomeProcedureOuTextoSQL;
                sqlCommand.CommandTimeout = 7200; //Em Segundos --- 2Horas.

                foreach (SqlParameter sqlParameter in sqlParameterCollection)
                {
                    sqlCommand.Parameters.Add(new SqlParameter(sqlParameter.ParameterName, sqlParameter.Value)).Direction = ParameterDirection.Input;
                }
                sqlCommand.Parameters.Add("@Retorno", SqlDbType.Int).Direction = ParameterDirection.Output;
                SqlParameter retorno = sqlCommand.Parameters.Add("@Retorno", SqlDbType.Int);
                retorno.Direction = ParameterDirection.ReturnValue;

                SqlDataReader dr = sqlCommand.ExecuteReader();

                existe = Convert.ToInt32(sqlCommand.Parameters["@Retorno"].Value.ToString());

                return existe;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public Typ PegaItem<Typ>(DataRow registro)
        {
            Type temp = typeof(Typ);
            Typ obj = Activator.CreateInstance<Typ>();

            foreach (DataColumn coluna in registro.Table.Columns)
            {
                foreach (PropertyInfo propriedade in temp.GetProperties())
                {
                    if (propriedade.Name == coluna.ColumnName)
                    {
                        if (registro[coluna.ColumnName] != DBNull.Value)
                        {
                            propriedade.SetValue(obj, registro[coluna.ColumnName], null);
                        }

                    }
                    else
                    {
                        continue;
                    }
                }
            }
            return obj;
        }

        public List<Typ> ConverteParaLista<Typ>(DataTable dt)
        {
            List<Typ> lst = new List<Typ>();
            foreach (DataRow row in dt.Rows)
            {
                Typ item = PegaItem<Typ>(row);
                lst.Add(item);
            }
            return lst;
        }



    }
}
