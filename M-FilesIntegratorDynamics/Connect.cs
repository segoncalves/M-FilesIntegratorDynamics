using CrmEarlyBound;
using Dapper;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;

namespace Integrador_MFD.Dynamics
{
    public class Connect
    {
        #region Class Level Members
        private IOrganizationService _orgService;

        #endregion Class Level Members

        #region Private Methods

        /// <summary>
        /// Verifies if a connection string is valid for Microsoft Dynamics CRM.
        /// </summary>
        /// <returns>True for a valid string, otherwise False.</returns>
        private static Boolean IsValidConnectionString(String connectionString)
        {
            // At a minimum, a connection string must contain one of these arguments.
            if (connectionString.Contains("Url=") ||
                connectionString.Contains("Server=") ||
                connectionString.Contains("ServiceUri="))
                return true;

            return false;
        }

        private static String GetServiceConfiguration()
        {
            // Get available connection strings from app.config.
            int count = ConfigurationManager.ConnectionStrings.Count;

            // Create a filter list of connection strings so that we have a list of valid
            // connection strings for Microsoft Dynamics CRM only.
            List<KeyValuePair<String, String>> filteredConnectionStrings =
                new List<KeyValuePair<String, String>>();

            for (int a = 0; a < count; a++)
            {
                if (IsValidConnectionString(ConfigurationManager.ConnectionStrings[a].ConnectionString))
                    filteredConnectionStrings.Add
                        (new KeyValuePair<string, string>
                            (ConfigurationManager.ConnectionStrings[a].Name,
                            ConfigurationManager.ConnectionStrings[a].ConnectionString));
            }

            // No valid connections strings found. Write out and error message.
            if (filteredConnectionStrings.Count == 0)
            {
                Console.WriteLine("An app.config file containing at least one valid Microsoft Dynamics CRM " +
                    "connection string configuration must exist in the run-time folder.");
                Console.WriteLine("\nThere are several commented out example connection strings in " +
                    "the provided app.config file. Uncomment one of them and modify the string according " +
                    "to your Microsoft Dynamics CRM installation. Then re-run the sample.");
                return null;
            }

            // If one valid connection string is found, use that.
            if (filteredConnectionStrings.Count == 1)
            {
                return filteredConnectionStrings[0].Value;
            }

            // If more than one valid connection string is found, let the user decide which to use.
            if (filteredConnectionStrings.Count > 1)
            {
                Console.WriteLine("The following connections are available:");
                Console.WriteLine("------------------------------------------------");

                for (int i = 0; i < filteredConnectionStrings.Count; i++)
                {
                    Console.Write("\n({0}) {1}\t",
                    i + 1, filteredConnectionStrings[i].Key);
                }

                Console.WriteLine();

                Console.Write("\nType the number of the connection to use (1-{0}) [{0}] : ",
                    filteredConnectionStrings.Count);
                String input = Console.ReadLine();
                int configNumber;
                if (input == String.Empty) input = filteredConnectionStrings.Count.ToString();
                if (!Int32.TryParse(input, out configNumber) || configNumber > count ||
                    configNumber == 0)
                {
                    Console.WriteLine("Option not valid.");
                    return null;
                }

                return filteredConnectionStrings[configNumber - 1].Value;
            }

            return null;
        }

        #endregion

        #region Public Methods

        public void Process()
        {
            List<DynamicsIntegratorPOC.Model.Lead> leads;

            DynamicsIntegratorPOC.Model.Lead lead = GetLastLead();

            leads = GetLeads(lead);

            if (leads != null)
                InsertDataBase(leads);
        }

        public List<DynamicsIntegratorPOC.Model.Lead> GetLeads(DynamicsIntegratorPOC.Model.Lead value)
        {
            string connectionString = GetServiceConfiguration();

            CrmServiceClient conn = new CrmServiceClient(connectionString);

            // Cast the proxy client to the IOrganizationService interface.
            _orgService = conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;

            // Obtain information about the logged on user from the web service.
            //Guid userid = ((WhoAmIResponse)_orgService.Execute(new WhoAmIRequest())).UserId;
            //SystemUser systemUser = (SystemUser)_orgService.Retrieve("systemuser", userid,
            //    new ColumnSet(new string[] { "firstname", "lastname" }));
            //Console.WriteLine("Logged on user is {0} {1}.", systemUser.FirstName, systemUser.LastName);

            // Retrieve the version of Microsoft Dynamics CRM.
            //RetrieveVersionRequest versionRequest = new RetrieveVersionRequest();
            //RetrieveVersionResponse versionResponse =
            //    (RetrieveVersionResponse)_orgService.Execute(versionRequest);
            //Console.WriteLine("Microsoft Dynamics CRM version {0}.", versionResponse.Version);

            // Retrieve attributes from a Lead (Prospect)
            ColumnSet cols = new ColumnSet(
                    "bz_tipo_prospect"
                    , "bz_razaosocial"
                    , "bz_cnpj"
                    , "fullname"
                    , "firstname"
                    , "bz_cpf"
                    , "bz_cpfcnpjrepresentante"
                    , "versionnumber"
                    , "leadid"
                    , "createdon"
            );

            QueryExpression opportunityProductsQuery;
            if (value == null)
            {
                opportunityProductsQuery = new QueryExpression
                {
                    EntityName = Lead.EntityLogicalName,
                    //ColumnSet = new ColumnSet("fullname", "firstname"),
                    ColumnSet = cols,
                };
            }
            else
            {
                opportunityProductsQuery = new QueryExpression
                {
                    EntityName = Lead.EntityLogicalName,
                    ColumnSet = cols,
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "createdon",
                                Operator = ConditionOperator.GreaterThan,
                                Values = { value.createdon }
                            }
                        }
                    }
                };
            }

            var ret = _orgService.RetrieveMultiple(opportunityProductsQuery).Entities;
            Console.WriteLine("Consumidos {0} Leads do Dynamics", ret.Count.ToString());

            if (ret.Count == 0)
                return null;

            List<DynamicsIntegratorPOC.Model.Lead> leads = new List<DynamicsIntegratorPOC.Model.Lead>();

            foreach (var item in ret)
            {
                DynamicsIntegratorPOC.Model.Lead lead = new DynamicsIntegratorPOC.Model.Lead();

                lead.leadid = item.Id.ToString("D");

                item.TryGetAttributeValue("fullname", out string fullname);
                if (fullname != null)
                    lead.fullname = fullname;

                item.TryGetAttributeValue("bz_tipo_prospect", out Microsoft.Xrm.Sdk.OptionSetValue bz_tipo_prospect);
                if (bz_tipo_prospect != null)
                    lead.bz_tipo_prospect = bz_tipo_prospect.Value;

                item.TryGetAttributeValue("bz_razaosocial", out string bz_razaosocial);
                if (bz_razaosocial != null)
                    lead.bz_razaosocial = bz_razaosocial;

                item.TryGetAttributeValue("bz_cnpj", out string bz_cnpj);
                if (bz_cnpj != null)
                    lead.bz_cnpj = bz_cnpj;

                item.TryGetAttributeValue("bz_cpf", out string bz_cpf);
                if (bz_cpf != null)
                    lead.bz_cpf = bz_cpf;

                item.TryGetAttributeValue("bz_cpfcnpjrepresentante", out string bz_cpfcnpjrepresentante);
                if (bz_cpfcnpjrepresentante != null)
                    lead.bz_cpfcnpjrepresentante = bz_cpfcnpjrepresentante;

                item.TryGetAttributeValue("versionnumber", out int versionnumber);
                if (versionnumber > 0)
                    lead.versionnumber = (versionnumber);

                item.TryGetAttributeValue("createdon", out DateTime createdon);
                if (createdon != null)
                    lead.createdon = (createdon);

                if (!ExistLead(lead.leadid))
                    leads.Add(lead);
            }

            Console.WriteLine("Organizados {0} Leads para inserção no banco", leads.Count.ToString());

            cols = null;

            _orgService = null;
            conn = null;

            return leads;
        }

        private int InsertDataBase(List<DynamicsIntegratorPOC.Model.Lead> leads)
        {
            string connectionString = "Data Source = ; Initial Catalog = ; Persist Security Info = True; User ID = ; Password = ";
            string SQL = "INSERT INTO [dbo].[Lead] VALUES (@leadid, @fullname , @bz_tipo_prospect, @bz_razaosocial, @bz_cnpj, @bz_cpf, @bz_cpfcnpjrepresentante, @versionnumber, @createdon)";
            int ret = 0;

            Console.WriteLine("Inserindo Leads no banco de dados");

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                ret = connection.Execute(SQL, leads);
            }

            Console.WriteLine("Inseridos {0} Leads no banco de dados", ret.ToString());

            return ret;
        }

        private DynamicsIntegratorPOC.Model.Lead GetLastLead()
        {
            //TODO: DON'T FORGET TO REMOVE
            string connectionString = "Data Source = ; Initial Catalog = ; Persist Security Info = True; User ID = ; Password = ";
            string SQL = "GetLastLead";
            DynamicsIntegratorPOC.Model.Lead lead;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                lead = connection.Query<DynamicsIntegratorPOC.Model.Lead>(SQL,
                    commandType: System.Data.CommandType.StoredProcedure).FirstOrDefault();
            }

            return lead;
        }

        private bool ExistLead(string leadid)
        {
            //TODO: DON'T FORGET TO REMOVE
            string connectionString = "Data Source = ; Initial Catalog = ; Persist Security Info = True; User ID = ; Password = ";
            string SQL = "ExistLead";
            int lead;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                lead = (int)connection.ExecuteScalar(SQL,
                    commandType: System.Data.CommandType.StoredProcedure,
                    param: new { @leadid = leadid});
            }

            return lead > 0;
        }

        #endregion
    }
}